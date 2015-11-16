VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpedidos 
   Caption         =   "Menu de pedidos sugeridos por tienda"
   ClientHeight    =   8595
   ClientLeft      =   255
   ClientTop       =   435
   ClientWidth     =   11880
   Icon            =   "frmpedidos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame frageped 
      Caption         =   "Generar pedido"
      Height          =   2415
      Left            =   1680
      TabIndex        =   37
      Top             =   2880
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton cmdgencan 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   4080
         Picture         =   "frmpedidos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdgenera 
         Caption         =   "Generar"
         Height          =   495
         Left            =   2040
         Picture         =   "frmpedidos.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cmbProv 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   41
         Text            =   "Combo2"
         Top             =   1200
         Width           =   5295
      End
      Begin VB.ComboBox cmbtienda 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   40
         Text            =   "Combo1"
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label lbletiqueta 
         Caption         =   "Proveedor"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbletiqueta 
         Caption         =   "Tienda"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc AdoDbf 
      Height          =   330
      Left            =   3960
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
   Begin VB.Frame fraPedsug 
      Height          =   5175
      Left            =   120
      TabIndex        =   28
      Top             =   2040
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton cmdRpt 
         Caption         =   "&Prod. Inactivos"
         Height          =   495
         Left            =   9720
         Picture         =   "frmpedidos.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Regresar a menu de pedidos sugeridos"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdImpReg 
         Caption         =   "&Regresar"
         Height          =   495
         Left            =   9720
         Picture         =   "frmpedidos.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Regresar a menu de pedidos sugeridos"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdImpImp 
         Caption         =   "&Importar"
         Height          =   495
         Left            =   9720
         Picture         =   "frmpedidos.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Importarel pedido sugerido"
         Top             =   720
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbgrdPedsuf 
         Bindings        =   "frmpedidos.frx":0E2C
         Height          =   4455
         Left            =   360
         TabIndex        =   32
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1.5
         RowHeight       =   17
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
         Caption         =   "PROVEEDORES ENVIADOS PARA PEDIDOS SUGERIDOS"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "claprove"
            Caption         =   "CLAVE PROVEEDOR"
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
            DataField       =   "consec"
            Caption         =   "    CVE. PROD."
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
            DataField       =   "cajas"
            Caption         =   "       CAJAS"
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
         BeginProperty Column03 
            DataField       =   "piezas"
            Caption         =   "        PIEZAS"
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
            DataField       =   "fecha"
            Caption         =   "        FECHA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/mm/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "pedido"
            Caption         =   "PED. SUGERIDO"
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
            Locked          =   -1  'True
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1934.929
            EndProperty
         EndProperty
      End
      Begin VB.Label lbpedidos 
         Caption         =   "0"
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
         Left            =   3000
         TabIndex        =   34
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL DE PEDIDOS :"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   4800
         Width           =   2175
      End
   End
   Begin MSAdodcLib.Adodc AdoPedidos 
      Height          =   330
      Left            =   1320
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
   Begin VB.PictureBox PicInf 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11820
      TabIndex        =   16
      Top             =   7650
      Width           =   11880
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   5
         Left            =   5040
         Picture         =   "frmpedidos.frx":0E41
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   90
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   4
         Left            =   4200
         Picture         =   "frmpedidos.frx":1373
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Buscar clave del pedido"
         Top             =   90
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   3
         Left            =   3360
         Picture         =   "frmpedidos.frx":146D
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ir al ultimo"
         Top             =   90
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   2
         Left            =   2520
         Picture         =   "frmpedidos.frx":15DF
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Ir al siguiente"
         Top             =   90
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   1
         Left            =   1680
         Picture         =   "frmpedidos.frx":1751
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ir al anterior"
         Top             =   90
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   0
         Left            =   840
         Picture         =   "frmpedidos.frx":18C3
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ir al primero"
         Top             =   90
         Width           =   600
      End
      Begin VB.PictureBox cRpt 
         Height          =   480
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   44
         Top             =   0
         Width           =   1200
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
         TabIndex        =   17
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame fradescripcion 
      Height          =   972
      Left            =   150
      TabIndex        =   10
      Top             =   960
      Width           =   11655
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   0
         Left            =   8280
         TabIndex        =   35
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   65077251
         CurrentDate     =   37257
      End
      Begin MSComDlg.CommonDialog Cmdlg 
         Left            =   2040
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   1
         Left            =   10080
         TabIndex        =   36
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   65077251
         CurrentDate     =   37257
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Fecha Inicial"
         Height          =   255
         Index           =   0
         Left            =   8160
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Fecha Final"
         Height          =   255
         Index           =   1
         Left            =   9960
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSucur 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label lblSuc 
         Caption         =   "Sucursal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblProve 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label lblProv 
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame cmdDespla 
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Exportar"
         Height          =   400
         Index           =   7
         Left            =   3360
         Picture         =   "frmpedidos.frx":1A35
         TabIndex        =   26
         ToolTipText     =   "Exporta los pedidos sugeridos en el rango especificado"
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Cancelar"
         Height          =   400
         Index           =   6
         Left            =   2280
         Picture         =   "frmpedidos.frx":1D3F
         TabIndex        =   5
         ToolTipText     =   "Cancela el pedido seleccionado"
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Importar"
         Height          =   375
         Index           =   4
         Left            =   4440
         Picture         =   "frmpedidos.frx":2049
         TabIndex        =   4
         ToolTipText     =   "Importa pedidos enviados por tiendas para su abastecimiento"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbOrden 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Regresar"
         Height          =   400
         Index           =   5
         Left            =   10200
         Picture         =   "frmpedidos.frx":2353
         TabIndex        =   7
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   3
         Left            =   7080
         Picture         =   "frmpedidos.frx":265D
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Recibir pedido y afectar inventario"
         Top             =   240
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Gen. Pedido"
         Height          =   375
         Index           =   2
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Confirmar pedido y prepararlo para recibirlo"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Modificar"
         Height          =   400
         Index           =   1
         Left            =   1200
         Picture         =   "frmpedidos.frx":2967
         TabIndex        =   2
         ToolTipText     =   "Modificar pedido capturado"
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Nuevo"
         Height          =   400
         Index           =   0
         Left            =   120
         Picture         =   "frmpedidos.frx":2C71
         TabIndex        =   1
         ToolTipText     =   "Capturar un nuevo pedido"
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label lblFiltro 
         Alignment       =   2  'Center
         Caption         =   "Filtro"
         Height          =   255
         Left            =   6840
         TabIndex        =   9
         Top             =   120
         Width           =   2895
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdPed 
      Bindings        =   "frmpedidos.frx":2F7B
      Height          =   5175
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9128
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "p_pedido"
         Caption         =   "FOLIO"
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
         DataField       =   "p_proveedor"
         Caption         =   "CVE. PROV."
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
         DataField       =   "p_sucursal"
         Caption         =   "CVE. SUC."
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
      BeginProperty Column03 
         DataField       =   "p_fecped"
         Caption         =   "FECHA  ELABORACION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "p_fecconfirma"
         Caption         =   "        FEC. DE CONF."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "p_fecentreal"
         Caption         =   "  FEC. DE RECEPCION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "p_notaentrada"
         Caption         =   " NOTA ENTRADA"
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
      BeginProperty Column07 
         DataField       =   "p_cancelado"
         Caption         =   "CANC."
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
         DataField       =   "p_sugerido"
         Caption         =   "SUG."
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
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   555.024
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   27
      Top             =   8265
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
End
Attribute VB_Name = "frmpedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private ccadconex As String
Private cCond As String     'Condicion del filtro del grid
Private ccondrpt As String  'Condicion del filtro del rpt
'Private rstSucProv As ADODB.Recordset
Private ntext As Integer
Private cFecha As String
Private cFecharpt As String

Private rstSucProv As ADODB.Recordset
Private cruta As String
Private cArch As String


Private Sub AdoPedidos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
lblProve.Caption = AdoPedidos.Recordset!NOMPROVE
lblSucur.Caption = AdoPedidos.Recordset!tidescrip
End Sub

Private Sub cmdgencan_Click()
    Me.frageped.Visible = False
End Sub

Private Sub cmdgenera_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
prov = Mid(cmbProv.Text, InStr(1, cmbProv.Text, "|") + 1)
sucu = Trim(Mid(cmbtienda.Text, InStr(1, cmbtienda.Text, "|") + 1))
rs.Open "SELECT p_pedido FROM pedidos WHERE p_sucursal = " & Trim(sucu) & " and p_proveedor = '" & prov & "' and NOT p_pedproveedor is null ORDER BY p_fecped DESC", cn, adOpenDynamic, adLockOptimistic, adCmdText
Pedido = "G" & rs!p_Pedido
MsgBox "INSERT INTO pedidos(p_pedido,p_proveedor,p_fecped,p_sucursal) VALUES ('" & Pedido & "','" & prov & "','" & date & "','" & sucu & "')"
cn.Execute "INSERT INTO pedidos(p_pedido,p_proveedor,p_fecped,p_sucursal) VALUES ('" & Pedido & "','" & prov & "','" & date & "','" & sucu & "')"
End Sub

Private Sub dtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtpFecha_LostFocus(Index As Integer)
On Error GoTo Error:
 
 cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
 cFecha = " AND (month(p_fecped) >= " & Month(dtpFecha(0).Value) & " and (day(p_fecped) > = " & Day(dtpFecha(0).Value) & cOper & " (day(p_fecped)<= " & Day(dtpFecha(1).Value) & " and month(p_fecped)<= " & Month(dtpFecha(1).Value) & ")) and year(p_fecped)>= " & Year(dtpFecha(0).Value) & " and year(p_fecped)<= " & Year(dtpFecha(1).Value) & ")"

 AdoPedidos.RecordSource = "SELECT * FROM Pedidos,CatProv,Cattienda WHERE " & cCond & cFecha & " ORDER BY p_fecped"
 AdoPedidos.Refresh
 lblInfo.Caption = "Numero de Pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
 For N = 0 To 5   'Si esta vacio el recordset desactivo las opciones
   Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 Next
 cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 cmdOpcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 Exit Sub
Error:
  MsgBox Err.Description
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
On Error GoTo Error:
Select Case cmbOrden.ListIndex
Case 0 'Todos
    cCond = "PEDIDOS.p_proveedor = CATPROV.prove AND PEDIDOS.p_sucursal = CATTIENDA.ticlave AND P_proveedor <> 'ABA'"
    ccondrpt = "{PEDIDOS.P_pedido} <> ''"
Case 1 'Pendientes por Confirmar
    cCond = "P_situacion = 0 AND PEDIDOS.p_proveedor = CATPROV.prove AND PEDIDOS.p_sucursal = CATTIENDA.ticlave AND P_proveedor <> 'ABA'"
    ccondrpt = "{PEDIDOS.P_situacion} = 0"
Case 2 'Pendientes por recibir
    cCond = "P_situacion = 1  AND P_recibido = 0 AND PEDIDOS.p_proveedor = CATPROV.prove AND PEDIDOS.p_sucursal = CATTIENDA.ticlave AND P_proveedor <> 'ABA'"
    ccondrpt = "{PEDIDOS.P_situacion} = 1  AND {PEDIDOS.P_recibido} = 0"
Case 3 'Recibidos
    cCond = "P_recibido = 1 AND PEDIDOS.p_proveedor = CATPROV.prove AND PEDIDOS.p_sucursal = CATTIENDA.ticlave AND P_proveedor <> 'ABA'"
    ccondrpt = "{PEDIDOS.P_recibido} = 1"
Case 4 'Sugeridos
    cCond = "p_sugerido = 1 AND PEDIDOS.p_proveedor = CATPROV.prove AND PEDIDOS.p_sucursal = CATTIENDA.ticlave AND P_proveedor <> 'ABA'"
    ccondrpt = "{PEDIDOS.P_sugerido} = 1"
End Select

cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
cFecha = " AND (month(p_fecped) >= " & Month(dtpFecha(0).Value) & " and (day(p_fecped) > = " & Day(dtpFecha(0).Value) & cOper & " (day(p_fecped)<= " & Day(dtpFecha(1).Value) & " and month(p_fecped)<= " & Month(dtpFecha(1).Value) & ")) and year(p_fecped)>= " & Year(dtpFecha(0).Value) & " and year(p_fecped)<= " & Year(dtpFecha(1).Value) & ")"

AdoPedidos.RecordSource = "SELECT * FROM Pedidos,CatProv,CatTienda WHERE " & cCond & cFecha
AdoPedidos.Refresh
For N = 0 To 5
    Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
Next
cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
cmdOpcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
'cmdOpcion(4).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
lblInfo.Caption = "Numero de pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdImpImp_Click()
   ImpPedsug
End Sub

Private Sub cmdImpReg_Click()
  Me.fraPedsug.Visible = False
End Sub

Private Sub cmdMoverse_Click(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0  'Primer registro
    AdoPedidos.Recordset.MoveFirst
Case 1  'Anterior
    AdoPedidos.Recordset.MovePrevious
    If AdoPedidos.Recordset.BOF Then AdoPedidos.Recordset.MoveFirst
Case 2  'Siguiente
    AdoPedidos.Recordset.MoveNext
    If AdoPedidos.Recordset.EOF Then AdoPedidos.Recordset.MoveLast
Case 3  'Ultimo
    AdoPedidos.Recordset.MoveLast
Case 4  'Buscar Clave de pedido
    cCve = InputBox("Introduzca la clave del pedido a buscar", "Introducir clave")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    Antes = dbgrdPed.Bookmark
    AdoPedidos.Recordset.MoveFirst
    AdoPedidos.Recordset.Find "p_pedido = '" & Trim(cCve) & "'"
    If AdoPedidos.Recordset.EOF Then
        MsgBox "LA CLAVE " & cCve & " NO SE ENCUENTRA EN LOS PEDIDOS " + IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text), vbExclamation
        dbgrdPed.Bookmark = Antes
    End If
Case 5
    cMensaje = stb1.SimpleText
    stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
    stb1.Refresh
    crpt.ReportFileName = App.Path & "\Pedidos.rpt"
    crpt.WindowTitle = "Reporte de pedidos sugeridos"
    crpt.Formulas(1) = "PEDIDO = 'LISTADO DE PEDIDOS SUGERIDOS DEL " & Trim(dtpFecha(0).Value) & " AL " & Trim(dtpFecha(1).Value) & " '"
    crpt.Connect = cCadConex
    crpt.SQLQuery = "SELECT PEDIDOS.p_pedido, PEDIDOS.p_proveedor, PEDIDOS.p_fecped, PEDIDOS.p_recibido, PEDIDOS.p_cancelado, PEDIDOS.P_sugerido, " & _
                            "CATPROV.NOMPROVE, DETALLEFACTURA.df_cantsol, DETALLEFACTURA.df_cantsolp, " & _
                            "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & Chr(13) & _
                    "FROM pitico.dbo.PEDIDOS PEDIDOS, " & _
                          "pitico.dbo.CATPROV CATPROV, " & _
                          "pitico.dbo.DETALLEFACTURA DETALLEFACTURA, " & _
                          "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                    "WHERE PEDIDOS.p_proveedor = CATPROV.PROVE AND " & _
                           "PEDIDOS.p_pedido = DETALLEFACTURA.df_pedido AND " & _
                           "DETALLEFACTURA.df_prod = TFPRODUC.CONSEC AND " & _
                           "PEDIDOS.p_fecped >= '" & Format(dtpFecha(0).Value, "yyyy-dd-mm") & "' AND PEDIDOS.p_fecped <= '" & Format(DateAdd("d", 1, dtpFecha(1).Value)) & "' AND " & _
                           "PEDIDOS.p_recibido = 0 AND PEDIDOS.p_cancelado = 0 AND " & _
                           "PEDIDOS.P_sugerido = 1 " & Chr(13) & _
                    "ORDER BY CATPROV.NOMPROVE ASC, PEDIDOS.p_pedido ASC, TFPRODUC.DESCRIPC ASC"
    'MsgBox cRpt.SQLQuery
    crpt.Action = 1
    stb1.SimpleText = cMensaje
    stb1.Refresh
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub



Private Sub cmdOpcion_Click(Index As Integer)
On Error GoTo Error:
Select Case Index
  Case 0  'Nuevo pedido
       cModo = "CAPTURARPEDIDO"
       nOp = 1
       frmCaptPed.Caption = "Nuevo pedido sugerido"
       frmCaptPed.Show
  Case 1  'Modificar pedido
       cModo = "CAPTURARPEDIDO"
       nOp = 0
       frmCaptPed.Caption = "Modificar pedido sugerido"
       frmCaptPed.Show
       SendKeys frmpedidos.dbgrdPed.Columns(0).Text
       SendKeys "{TAB}"
  Case 2  'Confirmar pedido
       frageped.Visible = True
       Dim rs As ADODB.Recordset
       Set rs = New ADODB.Recordset
       rs.Open "SELECT * FROM catprov WHERE activo = 1", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
       cmbProv.Clear
       While Not rs.EOF
            cmbProv.AddItem rs!NOMPROVE & "  |" & rs!prove
            rs.MoveNext
       Wend
       rs.Close
       rs.Open "SELECT * FROM cattienda", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
       cmbtienda.Clear
       While Not rs.EOF
            cmbtienda.AddItem rs!tidescrip & "  | " & rs!ticlave
            rs.MoveNext
       Wend
       rs.Close
       Set rs = Nothing
       Call Generasug
  Case 3  'Recibir pedido
       cModo = "RECIBIRPEDIDO"
       nOp = 0
       frmCaptPed.Caption = "Modificar pedido"
       frmCaptPed.Show
       SendKeys frmpedidos.dbgrdPed.Columns(0).Text
       SendKeys "{TAB}"
  Case 4  'Importar pedidos de tabla Dbf de tiendas a Sql Server tabla PEDIDOS y DETALLEFACTURAS
       VerPedSug
  Case 5  'Salir del modulo de pedidos
       Unload Me
  Case 6
       If MsgBox("REALMENTE DESEAS CANCELAR EL PEDIDO CON FOLIO: " & AdoPedidos.Recordset!p_Pedido & Chr(13) & _
          " DEL PROVEEDOR " & lblProve.Caption, vbInformation + vbYesNo) = vbYes Then
          AdoPedidos.Recordset!P_CANCELADO = 1
          AdoPedidos.Recordset.Update
       End If
  Case 7
       ExpPedsug
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Generasug()

End Sub

Private Sub cmdopcion_GotFocus(Index As Integer)
If Index = 0 Then Unload frmAreaRecibo
End Sub

Private Sub cmdRpt_Click()
Dim rs As ADODB.Recordset
 AdoDbf.CommandType = adCmdText
 AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
 AdoDbf.RecordSource = "SELECT DISTINCT CLAPROVE ,CONSEC,CAJAS,PIEZAS,FECHA,NUM_PED FROM " & cArch & " ORDER BY CLAPROVE,CONSEC"
 AdoDbf.Refresh
 Set rs = New ADODB.Recordset
 lInd = False
 While Not AdoDbf.Recordset.EOF
    rs.Open "SELECT * FROM TFPRODUC WHERE CONSEC = '" & CStr(AdoDbf.Recordset!CONSEC) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rs.BOF And rs.EOF Then
       If lInd = False Then
          encab
          lInd = True
       End If
       Print #1, AdoDbf.Recordset!CONSEC & " " & "  ERROR:NO EXISTE EN EL CATALOGO DE PRODUCTOS"
    ElseIf Not rs!activo Then
       If lInd = False Then
          encab
          lInd = True
       End If
       Print #1, rs!CONSEC & " " & rs!descripc & " " & Str(rs!PAQUETES) & " x " & CStr(rs!CONTENID) & " " & rs!medida & "  ERROR:MARCADO COMO INACTIIVO"
    End If
    rs.Close
    AdoDbf.Recordset.MoveNext
 Wend
 Close #1
 If lInd Then
    Handle = Shell("NOTEPAD " & App.Path & "\ERR" & UCase(Mid(cArch, 4, 3)) & ".TXT", 1)
 Else
    MsgBox "NO EXISTEN ERRORES EN EL PEDIDO SUGERIDO", vbInformation
 End If

End Sub

Private Sub encab()
 Open App.Path & "\ERR" & UCase(Mid(cArch, 4, 3)) & ".TXT" For Output As #1 ' Abre el archivo para operaciones de salida.
 Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
 Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
 Print #1, "ERROR ENCONTRADO EN EL PEDIDO SUGERIDO " & cArch & " DEL " & UCase(Format(date, "long date"))  ' Escribe texto en el archivo.
 Print #1,   ' Imprime una línea en blanco en el archivo.
 Print #1, "=========================================================================================="
 Print #1, "CLAVE         DESCRIPCION                                   PRESENTACION"
 Print #1, "=========================================================================================="
End Sub

Private Sub dbgrdPed_DblClick()
  If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF) Then cmdOpcion_Click 1
End Sub


Private Sub DbGrdped_HeadClick(ByVal ColIndex As Integer)
  stb1.SimpleText = Space(65) + "Espere un momento ordenando Pedidos por " & dbgrdPed.Columns(ColIndex).Caption
  AdoPedidos.RecordSource = "SELECT * FROM Pedidos,cattienda,catprov WHERE " & cCond & cFecha & "ORDER BY " & dbgrdPed.Columns(ColIndex).DataField
  AdoPedidos.Refresh
  stb1.SimpleText = Space(85) + "Pedidos ordenandos por " & dbgrdPed.Columns(ColIndex).Caption
End Sub

Private Sub dbgrdPed_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 If dbgrdPed.SelBookmarks.Count > 0 Then dbgrdPed.SelBookmarks.Remove 0
 dbgrdPed.SelBookmarks.Add dbgrdPed.RowBookmark(dbgrdPed.Row)
End Sub

Private Sub Form_Activate()
  Forma = 1  'Se activa bandera de pedidos sugeridos
  
  If dtpFecha(0).Value = "01/01/02" Then dtpFecha(0).Value = Format(date, "dd/mm/yyyy")
  If dtpFecha(1).Value = "01/01/02" Then dtpFecha(1).Value = Format(date, "dd/mm/yyyy")
  
  cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
  cFecha = " AND (month(p_fecped) >= " & Month(dtpFecha(0).Value) & " and (day(p_fecped) > = " & Day(dtpFecha(0).Value) & cOper & " (day(p_fecped)<= " & Day(dtpFecha(1).Value) & " and month(p_fecped)<= " & Month(dtpFecha(1).Value) & ")) and year(p_fecped)>= " & Year(dtpFecha(0).Value) & " and year(p_fecped)<= " & Year(dtpFecha(1).Value) & ")"  'Cargo todos los pedidos
  
  AdoPedidos.ConnectionString = cCadConex
  AdoPedidos.CommandType = adCmdText
  AdoPedidos.RecordSource = "SELECT * FROM Pedidos, Catprov, Cattienda WHERE " & cCond & cFecha
  AdoPedidos.Refresh
  For N = 0 To 5
     Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  Next
  cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 ' cmdopcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  'cmdOpcion(4).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  lblInfo.Caption = "Numero de pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
  cmbOrden.AddItem "TODOS"
  cmbOrden.AddItem "PENDIENTES DE CONFIRMAR"
  cmbOrden.AddItem "PENDIENTES DE RECIBIR"
  cmbOrden.AddItem "RECIBIDOS"
  cmbOrden.AddItem "SUGERIDOS"
  cmbOrden.ListIndex = 0
  cCond = "PEDIDOS.p_proveedor = CATPROV.prove AND PEDIDOS.p_sucursal = CATTIENDA.ticlave AND P_proveedor <> 'ABA'"             ' Filtro por default todos los pedidos
  ccondrpt = "{PEDIDOS.p_pedido} <> ''"  ' Filtro por default del RPT
  cmdOpcion(4).Enabled = tipotienda = 1
  cmdOpcion(7).Enabled = tipotienda = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmAreaRecibo.Show
End Sub

Private Sub ExpPedsug()
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
On Error GoTo Error:
   cMenAnt = stb1.SimpleText
   Cmdlg.Filter = "Pedidos sugeridos (ped" & LCase(Mid(Trim(Mid(cSucursal, 3)), 1, 3)) & ".dbf) | PED" & Mid(Trim(Mid(cSucursal, 3)), 1, 3) & ".dbf"
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

   'If MsgBox("DESEAS LIMPIAR EL ARCHIVO A ENVIAR", vbQuestion + vbYesNo) = vbYes Then
        stb1.SimpleText = Space(65) & "Limpiando archivo " & cArch
        stb1.Refresh
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFile("P:\EsPedSug.DBF")
        f.Copy cRutArc, True
   'End If
   
   Set rsttemp = New ADODB.Recordset
   AdoDbf.CommandType = adCmdText
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch & " ORDER BY CLAPROVE"
   AdoDbf.Refresh
   'Genero el archivo de reporte
   Open App.Path & "\PedSugEn.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
   Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
   Print #1, "     PEDIDOS SUGERIDOS PREPARADOS PARA ENVIAR EL "; UCase(Format(date, "long date"))  ' Escribe texto en el archivo.
   Print #1,   ' Imprime una línea en blanco en el archivo.
   Print #1, "=========================================================================================="
   Print #1, "          PROVEEDOR                                       FECHA  PROD. CAJAS  PIEZAS"
   Print #1, "=========================================================================================="

   If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF) Then AdoPedidos.Recordset.MoveFirst
   nPed = 0
   While Not AdoPedidos.Recordset.EOF
      If AdoPedidos.Recordset!p_sugerido And Not AdoPedidos.Recordset!P_CANCELADO And Not AdoPedidos.Recordset!P_recibido Then
        rsttemp.Open "SELECT * FROM detallefactura WHERE df_pedido = '" & AdoPedidos.Recordset!p_Pedido & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
        nProd = 0: nCajas = 0: nPiezas = 0
        While Not rsttemp.EOF
          stb1.SimpleText = Space(75) & "Exportando producto con la clave: " & CStr(rsttemp!df_prod)
          stb1.Refresh
          AdoDbf.Recordset.AddNew
          AdoDbf.Recordset!claprove = Trim(AdoPedidos.Recordset!p_proveedor)
          AdoDbf.Recordset!CONSEC = Val(Trim(rsttemp!df_prod))
          AdoDbf.Recordset!cajas = IIf(Trim(rsttemp!df_cantsol) = "", 0, Val(rsttemp!df_cantsol))
          AdoDbf.Recordset!piezas = IIf(Trim(rsttemp!df_cantsolp) = "", 0, Val(rsttemp!df_cantsolp))
          AdoDbf.Recordset!fecha = Mid(AdoPedidos.Recordset!p_fecped, 1, 8)
          AdoDbf.Recordset!Num_ped = Trim(AdoPedidos.Recordset!p_Pedido)
          AdoDbf.Recordset.Update
          nProd = nProd + 1: nCajas = nCajas + Val(rsttemp!df_cantsol): nPiezas = nPiezas + Val(rsttemp!df_cantsolp)
          rsttemp.MoveNext
        Wend
        rsttemp.Close
        Print #1, AdoPedidos.Recordset!p_proveedor & Trim(AdoPedidos.Recordset!NOMPROVE) & String(50 - Len(Trim(AdoPedidos.Recordset!NOMPROVE)), " ") & " " & Format(AdoPedidos.Recordset!p_fecped, "dd/mm/yy") & " " & nProd; Spc(5); nCajas; Spc(5); nPiezas
        nPed = nPed + 1
      End If
      AdoPedidos.Recordset.MoveNext
   Wend
   AdoDbf.Recordset.Close
   stb1.SimpleText = cMenAnt
   stb1.Refresh
   MsgBox "SE GENERARON UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
   Print #1, "=========================================================================================="
   Print #1, "TOTAL DE PEDIDOS: "; nPed
   Close #1   ' Cierra el archivo de reporte
   Handle = Shell("NOTEPAD " & App.Path & "\PedSugEn.TXT", 1)
  Exit Sub
Error:
   MsgBox UCase(Err.Description), vbCritical
   StbMensajes.SimpleText = cMenAnt
End Sub


Private Sub VerPedSug() '(CveTie As String)
 On Error GoTo Error:
 
 MenAnt = stb1.SimpleText
 Cmdlg.FileName = ""
 Cmdlg.CancelError = True
 Cmdlg.DialogTitle = "Abrir archivo de pedidos sugeridos enviado por tienda"
 Cmdlg.Filter = "Archivos Visual Foxpro (*.dbf) | *.dbf"
 Cmdlg.ShowOpen
 cArch = Cmdlg.FileName
 If cArch = "" Or IsNull(cArch) Then
    MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
    Exit Sub
 End If
   
 'Siempre el nombre de archivo es de 6 Caracteres
 cruta = Mid(cArch, 1, Len(cArch) - 10)
 cArch = Mid(cArch, Len(cArch) - 9)
 
 AdoDbf.CommandType = adCmdText
 AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
 AdoDbf.RecordSource = "SELECT CLAPROVE AS CLAPROVE, SUM(CAJAS) CAJAS, SUM(PIEZAS) AS PIEZAS, MAX(FECHA) AS FECHA, MAX(NUM_PED) AS PEDIDO FROM " & cArch & " GROUP BY claprove"
 AdoDbf.Refresh
 lbpedidos.Caption = AdoDbf.Recordset.RecordCount
 lbpedidos.Refresh
 fraPedsug.Visible = True
Exit Sub
Error:
  If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
     MsgBox Err.Description
  End If
End Sub

Private Sub ImpPedsug()
Me.dbgrdPedsuf.Columns(1).Visible = True
 AdoDbf.CommandType = adCmdText
 'AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties='DSN=Visual FoxPro Tables;UID=;SourceDB=" & cRuta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=SPANISH;Null=Sí;Deleted=Sí;';Initial Catalog= " & cRuta
 AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
 AdoDbf.RecordSource = "SELECT DISTINCT CLAPROVE ,CONSEC,CAJAS,PIEZAS,FECHA,NUM_PED FROM " & cArch & " ORDER BY CLAPROVE,CONSEC"
 AdoDbf.Refresh
 
 If AdoDbf.Recordset.BOF And AdoDbf.Recordset.EOF Then
    MsgBox "EL ARCHIVO A IMPORTAR NO TIENE NINGUN PRODUCTO"
       Exit Sub
 ElseIf Format(AdoDbf.Recordset!fecha, "DD/MM/YY") = "01/01/00" Then
    MsgBox "EL ARCHIVO " & UCase(cArch) & " YA FUE IMPORTADO", vbExclamation
    Exit Sub
 End If
 stb1.SimpleText = Space(45) & "Buscando Folio consecutivo de pedido"
 stb1.Refresh
 'Agrego el detalle de Factura
 AdoDbf.Recordset.MoveFirst
 cProAnt = "": FolPed = "": nPed = 0
 
 If Month(date) < 10 Then
    mesP = "0" & Trim(Str(Month(date)))
 Else
    mesP = Trim(Str(Month(date)))
 End If
 If Day(date) < 10 Then
    DiaP = "0" & Trim(Str(Day(date)))
 Else
    DiaP = Trim(Str(Day(date)))
 End If
 
  Open App.Path & "\PED" & Mid(cArch, 4, 3) & "-" & mesP & DiaP & ".TXT" For Output As #1 ' Abre el archivo para operaciones de salida.
   Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
   Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
   Print #1, "PEDIDOS POR SUGERIDOS  RECIBIDOS EL "; UCase(Format(date, "long date")) ' Escribe texto en el archivo.
   Print #1,   ' Imprime una línea en blanco en el archivo.
   Print #1, "=========================================================================================="
   Print #1, "PEDIDO        FEC. ELAB.              PROVEEDOR"
   Print #1, "=========================================================================================="
   Print #1,   ' Imprime una línea en blanco en el archivo.
 While Not AdoDbf.Recordset.EOF
       'Agrego Pedido por proveedor
       If AdoDbf.Recordset!claprove <> cProAnt Then
          Set rsttemp = New ADODB.Recordset
          'Obtengo el folio del pedido de la tienda
          ' tomando las tres primeras letras de la tienda que envia el pedido
          rsttemp.Open "SELECT MAX (CAST(SUBSTRING(P_PEDIDO,4,7) AS INT)) As FolMay FROM [PEDIDOS] WHERE SUBSTRING(p_pedido,1,3) = '" & Mid(Trim(cArch), 4, 3) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
          If IsNull(rsttemp!FolMay) Then
             FolPed = UCase(Mid(Trim(cArch), 4, 3)) & "1"
          Else
             FolPed = UCase(Mid(Trim(cArch), 4, 3)) & Trim(Str(rsttemp!FolMay + 1))
          End If
          SUC = Pedsuc(cArch)
          
          If IsNull(SUC) Or Trim(SUC) = "" Then
              MsgBox "EL NOMBRE DEL ARCHIVO: " & cArch & "NO ESTA REGISTRADO EN EL SISTEMA" & Chr(13) _
              & "Y NO SE IMPORTARAN LOS DATOS" & Chr(13) & "FAVOR DE AVISAR AL ADMINISTRADOR DEL SISTEMA", vbCritical
              Exit Sub
          End If
          nPed = nPed + 1
          cn.Execute "INSERT INTO Pedidos(p_pedido,p_proveedor,p_fecped,p_sucursal,p_observaciones) SELECT Ped ='" & FolPed & "', Prove = '" & AdoDbf.Recordset!claprove & "', Fechasol = '" & AdoDbf.Recordset!fecha & "',Suc = '" & SUC & "' ,Ped = 'PEDIDO IMPORTADO DEL PEDIDO " & AdoDbf.Recordset!Num_ped & "'"
          'MsgBox "SE GENERO EL PEDIDO POR TIENDA CON FOLIO: " & FolPed & Chr(13) & " DEL PROVEEDOR CON CLAVE: " & AdoDbf.Recordset!claprove & Space(5) & "CON FECHA: " & AdoDbf.Recordset!Fecha, vbInformation
          Print #1, Space(2) & FolPed & "    " & date & "  " & Space(15) & AdoDbf.Recordset!claprove
       End If
       stb1.SimpleText = Space(75) & "Importando producto: " & AdoDbf.Recordset!CONSEC '& Space(3)  & AdoDbf.Recordset!Descripc
       stb1.Refresh
       rsttemp.Close
       rsttemp.Open "SELECT * FROM TFPRODUC WHERE CONSEC = '" & CStr(AdoDbf.Recordset!CONSEC) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
       lActivo = True
       If rsttemp.BOF And rsttemp.EOF Then
          lActivo = False
          MsgBox "EL PRODUCTO NO SE ENCUENTRA EN EL CATALOGO, ANOTE LOS DATOS E INFORME AL ADMINISTRADOR DEL SISTEMA" & Chr(13) & "PEDIDO:" & FolPed & Chr(13) & "PROVEEDOR: " & AdoDbf.Recordset!claprove & Chr(13) & "CLAVE PRODUCTO: " & CStr(AdoDbf.Recordset!CONSEC) & Chr(13) & "CAJAS: " & CStr(AdoDbf.Recordset!cajas) & Chr(13) & "PIEZAS: " & CStr(AdoDbf.Recordset!piezas), vbCritical
       ElseIf Not rsttemp!activo Then
          MsgBox "EL PRODUCTO " & rsttemp!descripc & Chr(13) & "ESTA INACTIVO, POR LO TANTO NO SE IMPORTARA" & Chr(13) & "CLAVE PRODUCTO: " & rsttemp!CONSEC & Chr(13) & "CAJAS: " & CStr(AdoDbf.Recordset!cajas) & Chr(13) & "PIEZAS: " & CStr(AdoDbf.Recordset!piezas), vbCritical
          lActivo = False
       End If
       'Hago esto y no un EXECUTE para que no truene en el caso de que no exista el producto en descprod o el costo sea NULO
       If lActivo = True Then
            rsttemp.Close
            rsttemp.Open "SELECT Costo As Pre FROM DESCPROD WHERE producto = '" & CStr(AdoDbf.Recordset!CONSEC) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
            nprecio = IIf(rsttemp.BOF And rsttemp.EOF, 0, rsttemp!pre)
            cn.Execute "INSERT INTO DetalleFactura(df_prod,df_pedido,df_cantsol,df_cantsolp,df_costo) VALUES ('" & AdoDbf.Recordset!CONSEC & "','" & FolPed & "'," & CStr(AdoDbf.Recordset!cajas) & "," & CStr(AdoDbf.Recordset!piezas) & "," & nprecio & ")"
       End If
       cProAnt = AdoDbf.Recordset!claprove
       AdoDbf.Recordset.MoveNext
 Wend
 'Actualizo el costo de los producto del detalle de factura porque el ultimo pedido no se actualiza
 'If FolPed <> "" Then cn.Execute "UPDATE DetalleFactura SET DetalleFactura.DF_COSTO = DESCPROD.COSTO FROM DetalleFactura, DESCPROD WHERE DESCPROD.PRODUCTO = DetalleFactura.Df_prod AND DetalleFactura.DF_PEDIDO = '" & FolPed & "'"
 MsgBox "SE GENERO UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
 'Pongo marca en el campo fecha de que el pedido ya fue importado
Print #1, "=========================================================================================="
Print #1, "TOTAL DE PEDIDOS: "; nPed
Close #1   ' Cierra el archivo de reporte
Handle = Shell("NOTEPAD " & App.Path & "\PED" & Mid(cArch, 4, 3) & "-" & mesP & DiaP & ".TXT", 1)
   
 Set cnFoxPro = New ADODB.Connection
 cnFoxPro.ConnectionString = "Provider=MSDASQL.1;DSN=Visual FoxPro Tables;SourceDB=" & cruta & ";SourceType=DBF;Exclusive=No;Initial Catalog= " & cruta
 cnFoxPro.Open
 cnFoxPro.Execute "UPDATE " & cArch & " SET FECHA = CTOD('01/01/2000')"

 stb1.SimpleText = MenAnt
 stb1.Refresh
 AdoPedidos.Refresh
Exit Sub
Error:
  If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
     MsgBox Err.Description
  End If
End Sub

