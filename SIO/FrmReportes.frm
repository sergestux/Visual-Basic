VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReport 
   Caption         =   "Menu  de reportes.."
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11415
   Icon            =   "FrmReportes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Palette         =   "FrmReportes.frx":0442
   ScaleHeight     =   8340
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1215
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   10455
      Begin VB.ComboBox cmbProd 
         Height          =   315
         ItemData        =   "FrmReportes.frx":0944
         Left            =   960
         List            =   "FrmReportes.frx":0946
         Sorted          =   -1  'True
         TabIndex        =   122
         Top             =   720
         Width           =   5055
      End
      Begin VB.ComboBox cmbProVta 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   121
         Top             =   240
         Width           =   5055
      End
      Begin MSMask.MaskEdBox mskHoraFin 
         Height          =   300
         Left            =   9840
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskhoraIni 
         Height          =   300
         Left            =   8640
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton btnsal 
         Caption         =   "Regresa&r"
         Height          =   450
         Index           =   5
         Left            =   9120
         Picture         =   "FrmReportes.frx":0948
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton btnreport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reporte"
         Height          =   500
         Left            =   9120
         Picture         =   "FrmReportes.frx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Presentación preeliminar del reporte"
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtfecini 
         Height          =   300
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   21364737
         CurrentDate     =   37178
      End
      Begin MSComCtl2.DTPicker dtfecfin 
         Height          =   300
         Left            =   6840
         TabIndex        =   10
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   21364737
         CurrentDate     =   37178
      End
      Begin VB.Label Label8 
         Caption         =   "Producto"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   123
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   119
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Hora Fin."
         Height          =   255
         Index           =   1
         Left            =   9840
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Hora Ini."
         Height          =   255
         Index           =   0
         Left            =   8640
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Final"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Inicial"
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin TabDlg.SSTab TabRep 
      Height          =   6240
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   11007
      _Version        =   393216
      Tabs            =   9
      Tab             =   8
      TabsPerRow      =   6
      TabHeight       =   617
      ForeColor       =   4194304
      MouseIcon       =   "FrmReportes.frx":0FEC
      TabCaption(0)   =   "&Créditos"
      TabPicture(0)   =   "FrmReportes.frx":1008
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Optcred(2)"
      Tab(0).Control(1)=   "Optcred(1)"
      Tab(0).Control(2)=   "Optcred(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Ped. x &tienda"
      TabPicture(1)   =   "FrmReportes.frx":114E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "OptPedido(4)"
      Tab(1).Control(1)=   "txtfecha"
      Tab(1).Control(2)=   "OptPedido(3)"
      Tab(1).Control(3)=   "cmbProv"
      Tab(1).Control(4)=   "cmdPedido"
      Tab(1).Control(5)=   "OptPedido(2)"
      Tab(1).Control(6)=   "OptPedido(1)"
      Tab(1).Control(7)=   "OptPedido(0)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Ped. por prov."
      TabPicture(2)   =   "FrmReportes.frx":12A8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbletiquetas(7)"
      Tab(2).Control(1)=   "lbletiquetas(6)"
      Tab(2).Control(2)=   "optPedpro(4)"
      Tab(2).Control(3)=   "optPedpro(3)"
      Tab(2).Control(4)=   "cmbProvPr"
      Tab(2).Control(5)=   "optPedpro(0)"
      Tab(2).Control(6)=   "optPedpro(1)"
      Tab(2).Control(7)=   "optPedpro(2)"
      Tab(2).Control(8)=   "cmdPedPro"
      Tab(2).Control(9)=   "optPedprove(1)"
      Tab(2).Control(10)=   "optPedprove(0)"
      Tab(2).Control(11)=   "fraPer"
      Tab(2).Control(12)=   "optPedpro(5)"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "&Inventario"
      TabPicture(3)   =   "FrmReportes.frx":13BA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Adoexcel"
      Tab(3).Control(1)=   "OptExis(8)"
      Tab(3).Control(2)=   "lstNoenv"
      Tab(3).Control(3)=   "OptExis(7)"
      Tab(3).Control(4)=   "OptExis(6)"
      Tab(3).Control(5)=   "fraexp"
      Tab(3).Control(6)=   "txtExis"
      Tab(3).Control(7)=   "cmbClas(0)"
      Tab(3).Control(8)=   "cmbClas(3)"
      Tab(3).Control(9)=   "cmbClas(2)"
      Tab(3).Control(10)=   "cmbClas(1)"
      Tab(3).Control(11)=   "OptExis(0)"
      Tab(3).Control(12)=   "OptExis(1)"
      Tab(3).Control(13)=   "OptExis(2)"
      Tab(3).Control(14)=   "OptExis(3)"
      Tab(3).Control(15)=   "OptExis(4)"
      Tab(3).Control(16)=   "cmbProved"
      Tab(3).Control(17)=   "OptExis(5)"
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "&Traslados"
      TabPicture(4)   =   "FrmReportes.frx":1554
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Rpt"
      Tab(4).Control(1)=   "fraTrasl"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "&Ventas"
      TabPicture(5)   =   "FrmReportes.frx":1776
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "OptVentas(7)"
      Tab(5).Control(1)=   "OptVentas(6)"
      Tab(5).Control(2)=   "OptVentas(5)"
      Tab(5).Control(3)=   "OptVentas(4)"
      Tab(5).Control(4)=   "OptVentas(3)"
      Tab(5).Control(5)=   "OptVentas(1)"
      Tab(5).Control(6)=   "OptVentas(2)"
      Tab(5).Control(7)=   "OptVentas(0)"
      Tab(5).Control(8)=   "FraVentas"
      Tab(5).ControlCount=   9
      TabCaption(6)   =   "&Productos"
      TabPicture(6)   =   "FrmReportes.frx":18A0
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "CmdDifPro"
      Tab(6).Control(1)=   "Opnalim(4)"
      Tab(6).Control(2)=   "Opnalim(3)"
      Tab(6).Control(3)=   "Opnalim(0)"
      Tab(6).Control(4)=   "Opnalim(1)"
      Tab(6).Control(5)=   "Opnalim(2)"
      Tab(6).Control(6)=   "Opnalim(5)"
      Tab(6).Control(7)=   "Opnalim(6)"
      Tab(6).Control(8)=   "cmbProvPro"
      Tab(6).ControlCount=   9
      TabCaption(7)   =   "&Proveedores"
      TabPicture(7)   =   "FrmReportes.frx":18BC
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Opprov(8)"
      Tab(7).Control(1)=   "Opprov(4)"
      Tab(7).Control(2)=   "Opprov(2)"
      Tab(7).Control(3)=   "Opprov(1)"
      Tab(7).Control(4)=   "Opprov(0)"
      Tab(7).Control(5)=   "Opprov(5)"
      Tab(7).Control(6)=   "Opprov(6)"
      Tab(7).Control(7)=   "Opprov(3)"
      Tab(7).Control(8)=   "Opprov(7)"
      Tab(7).Control(9)=   "cmbtipo"
      Tab(7).ControlCount=   10
      TabCaption(8)   =   "Agentes"
      TabPicture(8)   =   "FrmReportes.frx":18D8
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "fraPedidos"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      Begin VB.OptionButton OptVentas 
         Caption         =   "Desplazamiento cajas y piezas"
         Height          =   255
         Index           =   7
         Left            =   -68280
         TabIndex        =   127
         ToolTipText     =   "Ventas clasificadas por serie y pendientes de cobro"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.OptionButton OptVentas 
         Caption         =   "&Factor estacional"
         Height          =   255
         Index           =   6
         Left            =   -74280
         TabIndex        =   126
         Top             =   3720
         Width           =   3015
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "Rentabilidad por proveedor"
         Height          =   195
         Index           =   8
         Left            =   -73800
         TabIndex        =   125
         Top             =   4080
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc Adoexcel 
         Height          =   330
         Left            =   -68160
         Top             =   4200
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
         Enabled         =   0
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoExcel"
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
      Begin VB.OptionButton OptVentas 
         Caption         =   "Ley de pareto  80% - 20 %"
         Height          =   255
         Index           =   5
         Left            =   -74280
         TabIndex        =   124
         ToolTipText     =   "Ventas clasificadas por serie y pendientes de cobro"
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cmbProvPro 
         Height          =   315
         Left            =   -73320
         Sorted          =   -1  'True
         TabIndex        =   120
         Top             =   5040
         Width           =   5655
      End
      Begin VB.OptionButton Opnalim 
         Caption         =   "Lista de precios por proveedor"
         Height          =   195
         Index           =   6
         Left            =   -71400
         TabIndex        =   117
         Top             =   4320
         Width           =   2775
      End
      Begin VB.OptionButton Optcred 
         Caption         =   "Cartera de clientes con crédito en Puerto escondido"
         Height          =   495
         Index           =   2
         Left            =   -74400
         TabIndex        =   115
         Top             =   2400
         Width           =   2775
      End
      Begin VB.OptionButton Optcred 
         Caption         =   "Créditos otorgados en Puerto Escondido"
         Height          =   495
         Index           =   1
         Left            =   -74400
         TabIndex        =   114
         Top             =   3120
         Width           =   2175
      End
      Begin VB.OptionButton Optcred 
         Caption         =   "Cartera de clientes en esta Bodega"
         Height          =   495
         Index           =   0
         Left            =   -74400
         TabIndex        =   113
         Top             =   1680
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton OptVentas 
         Caption         =   "Ventas por agente y serie"
         Height          =   255
         Index           =   4
         Left            =   -74280
         TabIndex        =   111
         ToolTipText     =   "Ventas clasificadas por serie y pendientes de cobro"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.OptionButton optPedpro 
         Caption         =   "Acumulado de compras"
         Height          =   255
         Index           =   5
         Left            =   -73920
         TabIndex        =   110
         Top             =   4080
         Width           =   2655
      End
      Begin VB.OptionButton OptVentas 
         Caption         =   "&Comparativo de ventas año anterior"
         Height          =   255
         Index           =   3
         Left            =   -74280
         TabIndex        =   109
         Top             =   3240
         Width           =   3015
      End
      Begin VB.OptionButton OptVentas 
         Caption         =   "Ventas facturadas en esta bodega contado y credito "
         Height          =   255
         Index           =   1
         Left            =   -74280
         TabIndex        =   101
         ToolTipText     =   "Ventas clasificadas por serie y pendientes de cobro"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.OptionButton OptVentas 
         Caption         =   "Ventas facturadas todas las Bodegas de Mayoreo"
         Height          =   255
         Index           =   2
         Left            =   -74280
         TabIndex        =   100
         ToolTipText     =   "Ventas clasificadas por Bodegas de Mayoreo"
         Top             =   2280
         Width           =   4335
      End
      Begin VB.OptionButton OptVentas 
         Caption         =   "&Desplazamiento Bodegas de Mayoreo"
         Height          =   255
         Index           =   0
         Left            =   -74280
         TabIndex        =   94
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Frame FraVentas 
         BorderStyle     =   0  'None
         Caption         =   "Ventas mayoreo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74400
         TabIndex        =   83
         Top             =   3480
         Width           =   8775
         Begin VB.Frame FraOrden 
            Caption         =   "Tipo de orden"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   975
            Left            =   6840
            TabIndex        =   97
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
            Begin VB.OptionButton OptOrden 
               Caption         =   "Descendente"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   99
               Top             =   600
               Width           =   1335
            End
            Begin VB.OptionButton OptOrden 
               Caption         =   "Ascendente"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   98
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame FraTipVta 
            Caption         =   "Ordenado por"
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
            Height          =   2055
            Left            =   3720
            TabIndex        =   86
            Top             =   480
            Visible         =   0   'False
            Width           =   2535
            Begin VB.OptionButton OptTipVta 
               Caption         =   "Utilidad (Porcentaje %)"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   116
               Top             =   1800
               Width           =   2055
            End
            Begin VB.OptionButton OptTipVta 
               Caption         =   "Descripción"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   90
               Top             =   360
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.OptionButton OptTipVta 
               Caption         =   "Cajas"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   91
               Top             =   720
               Width           =   1935
            End
            Begin VB.OptionButton OptTipVta 
               Caption         =   "Importe"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   92
               Top             =   1080
               Width           =   2175
            End
            Begin VB.OptionButton OptTipVta 
               Caption         =   "Utilidad (Importe $)"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   93
               Top             =   1440
               Width           =   2055
            End
         End
         Begin VB.Frame fraSerie 
            Caption         =   "Sucursal"
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
            Height          =   1935
            Left            =   120
            TabIndex        =   84
            Top             =   480
            Visible         =   0   'False
            Width           =   2535
            Begin VB.OptionButton OptSerie 
               Caption         =   "28)  Istmo"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   118
               Top             =   1200
               Width           =   2295
            End
            Begin VB.OptionButton OptSerie 
               Caption         =   "Todas las sucursales"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   95
               Top             =   1560
               Value           =   -1  'True
               Width           =   2175
            End
            Begin VB.OptionButton OptSerie 
               Caption         =   "13)  Miguel Cabrera"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   85
               Top             =   240
               Width           =   2295
            End
            Begin VB.OptionButton OptSerie 
               Caption         =   "23)  Central de Abastos "
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   87
               Top             =   480
               Width           =   2295
            End
            Begin VB.OptionButton OptSerie 
               Caption         =   "55)  Puerto Escondido"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   88
               Top             =   720
               Width           =   2295
            End
            Begin VB.OptionButton OptSerie 
               Caption         =   "26)  Miahuatlan"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   89
               Top             =   960
               Width           =   2175
            End
         End
      End
      Begin VB.Frame fraPedidos 
         BorderStyle     =   0  'None
         Caption         =   "Reporte de ventas por Agente"
         ForeColor       =   &H80000002&
         Height          =   5055
         Left            =   1200
         TabIndex        =   78
         Top             =   960
         Width           =   8175
         Begin VB.OptionButton OptAgente 
            Caption         =   "Análisis de ruta"
            Height          =   255
            Index           =   10
            Left            =   3120
            TabIndex        =   112
            ToolTipText     =   "Análisis de ruta con equipos Handheld"
            Top             =   4680
            Width           =   2295
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Situación de  preventas"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   108
            ToolTipText     =   "Comprativo semanal de la misma ruta"
            Top             =   4680
            Width           =   2295
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Comparativo de preventas"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   107
            ToolTipText     =   "Comprativo semanal de la misma ruta"
            Top             =   4200
            Width           =   2295
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Utilidad por Agente"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   106
            ToolTipText     =   "Importe total de la Preventa agrupado por agente"
            Top             =   360
            Value           =   -1  'True
            Width           =   3015
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Utilidad por Cliente"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   105
            ToolTipText     =   "Facturas clasificadas por preventa"
            Top             =   3720
            Width           =   1935
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Creditos por Agente"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   104
            ToolTipText     =   "Facturas clasificadas por preventa"
            Top             =   2760
            Width           =   1815
         End
         Begin VB.ComboBox cmbCliente 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   103
            Top             =   3480
            Visible         =   0   'False
            Width           =   6015
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Facturado por Cliente"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   102
            ToolTipText     =   "Facturas clasificadas por preventa"
            Top             =   3240
            Width           =   2055
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Ventas detalladas clasificadas por precios normal, especial y por período"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   96
            ToolTipText     =   "Ventas detalladas clasificadas por precio normal o especial"
            Top             =   1800
            Width           =   5535
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Ventas clasificadas por preventa"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "Importe total de la Preventa agrupado por agente"
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox cmbAgentes 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   81
            Top             =   0
            Width           =   6135
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Facturas por preventa"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   80
            ToolTipText     =   "Facturas clasificadas por preventa"
            Top             =   2280
            Width           =   2175
         End
         Begin VB.OptionButton OptAgente 
            Caption         =   "Ventas clasificadas por precios normal, especial y por preventa"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   79
            ToolTipText     =   "Ventas clasificadas por precio normal o especial"
            Top             =   1320
            Width           =   5055
         End
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "Dias Inventario base ventas"
         Height          =   255
         Index           =   8
         Left            =   -73320
         TabIndex        =   76
         Top             =   4800
         Width           =   2895
      End
      Begin VB.OptionButton Opnalim 
         Caption         =   "Inactivados por Periodo"
         Height          =   195
         Index           =   5
         Left            =   -71400
         TabIndex        =   75
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ListBox lstNoenv 
         Height          =   2760
         ItemData        =   "FrmReportes.frx":18F4
         Left            =   -69720
         List            =   "FrmReportes.frx":18FB
         Style           =   1  'Checkbox
         TabIndex        =   74
         Top             =   1440
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "Existencias en CDC y no en tiendas"
         Height          =   255
         Index           =   7
         Left            =   -73320
         TabIndex        =   73
         Top             =   4320
         Width           =   2895
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "Activos"
         Height          =   195
         Index           =   4
         Left            =   -69000
         TabIndex        =   72
         Top             =   2520
         Width           =   1935
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "Foráneos"
         Height          =   195
         Index           =   2
         Left            =   -73800
         TabIndex        =   71
         Top             =   2880
         Width           =   2895
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "Locales"
         Height          =   195
         Index           =   1
         Left            =   -73800
         TabIndex        =   70
         Top             =   2280
         Width           =   2895
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "General  (Locales, foraneos, activos e inactivos)"
         Height          =   195
         Index           =   0
         Left            =   -73800
         TabIndex        =   69
         Top             =   1680
         Width           =   3975
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "Inactivos"
         Height          =   195
         Index           =   5
         Left            =   -69000
         TabIndex        =   68
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "Proveedores c/representantes"
         Height          =   195
         Index           =   6
         Left            =   -69000
         TabIndex        =   67
         Top             =   1680
         Width           =   2775
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "Clasificados por comprador (a)"
         Height          =   195
         Index           =   3
         Left            =   -73800
         TabIndex        =   66
         Top             =   3480
         Width           =   2895
      End
      Begin VB.OptionButton Opprov 
         Caption         =   "Por tipo"
         Height          =   195
         Index           =   7
         Left            =   -69000
         TabIndex        =   65
         Top             =   4200
         Width           =   1215
      End
      Begin VB.ComboBox cmbtipo 
         Height          =   315
         ItemData        =   "FrmReportes.frx":1909
         Left            =   -67320
         List            =   "FrmReportes.frx":190B
         TabIndex        =   64
         Top             =   4200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "Exportar existencias a Excel"
         Height          =   255
         Index           =   6
         Left            =   -73320
         TabIndex        =   63
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Frame fraexp 
         Height          =   735
         Left            =   -74880
         TabIndex        =   61
         Top             =   5160
         Width           =   9615
         Begin ComctlLib.ProgressBar ProBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   661
            _Version        =   327682
            Appearance      =   1
         End
      End
      Begin VB.TextBox txtExis 
         Height          =   375
         Left            =   -68520
         TabIndex        =   60
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbClas 
         Height          =   315
         Index           =   0
         Left            =   -70440
         TabIndex        =   59
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cmbClas 
         Height          =   315
         Index           =   3
         Left            =   -70920
         Sorted          =   -1  'True
         TabIndex        =   58
         Top             =   2400
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.ComboBox cmbClas 
         Height          =   315
         Index           =   2
         Left            =   -70920
         Sorted          =   -1  'True
         TabIndex        =   57
         Top             =   1440
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.ComboBox cmbClas 
         Height          =   315
         Index           =   1
         Left            =   -70920
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   1920
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "&Existencia"
         Height          =   375
         Index           =   0
         Left            =   -73320
         TabIndex        =   55
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "Existencia por familia"
         Height          =   375
         Index           =   1
         Left            =   -73320
         TabIndex        =   54
         Top             =   1920
         Width           =   2295
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "Existencia por linea"
         Height          =   375
         Index           =   2
         Left            =   -73320
         TabIndex        =   53
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "Existencia por departamento"
         Height          =   255
         Index           =   3
         Left            =   -73320
         TabIndex        =   52
         Top             =   2400
         Width           =   2415
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "Existencia por proveedor"
         Height          =   255
         Index           =   4
         Left            =   -73320
         TabIndex        =   51
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ComboBox cmbProved 
         Height          =   315
         Left            =   -70920
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   2880
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.OptionButton OptExis 
         Caption         =   "Existencia$"
         Height          =   255
         Index           =   5
         Left            =   -73320
         TabIndex        =   49
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Frame fraPer 
         Caption         =   "Período de entrega"
         Height          =   615
         Left            =   -70320
         TabIndex        =   45
         Top             =   3480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.OptionButton optper 
            Caption         =   "Mensual"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optper 
            Caption         =   "&Anual"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   47
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optper 
            Caption         =   "Diario"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.OptionButton optPedprove 
         Caption         =   "Eficiencia en el surtido"
         Height          =   255
         Index           =   0
         Left            =   -70560
         TabIndex        =   44
         Top             =   2640
         Width           =   2295
      End
      Begin VB.OptionButton optPedprove 
         Caption         =   "Llegadas de productos"
         Height          =   255
         Index           =   1
         Left            =   -70560
         TabIndex        =   43
         Top             =   3120
         Width           =   2295
      End
      Begin VB.CommandButton cmdPedPro 
         Caption         =   "Vista &Preeliminar"
         Height          =   495
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Ver reporte"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optPedpro 
         Caption         =   "&Pedidos recibidos"
         Height          =   255
         Index           =   2
         Left            =   -73920
         TabIndex        =   39
         Top             =   3120
         Width           =   1815
      End
      Begin VB.OptionButton optPedpro 
         Caption         =   "Pedidos pendientes de &recibir"
         Height          =   255
         Index           =   1
         Left            =   -73920
         TabIndex        =   38
         Top             =   2640
         Width           =   2775
      End
      Begin VB.OptionButton optPedpro 
         Caption         =   "Pedidos pendientes de &confirmar"
         Height          =   255
         Index           =   0
         Left            =   -73920
         TabIndex        =   37
         Top             =   2160
         Width           =   2895
      End
      Begin VB.ComboBox cmbProvPr 
         Height          =   315
         Left            =   -72840
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   1440
         Width           =   6375
      End
      Begin VB.OptionButton optPedpro 
         Caption         =   "Eficiencia en el &surtido global"
         Height          =   255
         Index           =   3
         Left            =   -73920
         TabIndex        =   35
         Top             =   3600
         Width           =   2655
      End
      Begin VB.OptionButton optPedpro 
         Caption         =   "Eficiencia en el &surtido por comprador"
         Height          =   195
         Index           =   4
         Left            =   -70560
         TabIndex        =   34
         Top             =   2160
         Width           =   3975
      End
      Begin VB.OptionButton OptPedido 
         Caption         =   "Pedidos pendientes por confirmar"
         Height          =   255
         Index           =   0
         Left            =   -73080
         TabIndex        =   33
         Top             =   2160
         Width           =   2775
      End
      Begin VB.OptionButton OptPedido 
         Caption         =   "Pedidos pendientes por recibir"
         Height          =   195
         Index           =   1
         Left            =   -73080
         TabIndex        =   32
         Top             =   2640
         Width           =   2775
      End
      Begin VB.OptionButton OptPedido 
         Caption         =   "Pedidos recibidos"
         Height          =   375
         Index           =   2
         Left            =   -73080
         TabIndex        =   31
         Top             =   3000
         Width           =   2775
      End
      Begin VB.CommandButton cmdPedido 
         Caption         =   "Vista &Preeliminar"
         Height          =   495
         Left            =   -70200
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Ver reporte"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.ComboBox cmbProv 
         Height          =   315
         Left            =   -72240
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   1560
         Width           =   6255
      End
      Begin VB.OptionButton OptPedido 
         Caption         =   "Eficiencia en surtido de pedidos"
         Height          =   255
         Index           =   3
         Left            =   -73080
         TabIndex        =   28
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtfecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70080
         TabIndex        =   27
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptPedido 
         Caption         =   "Total de entradas por tienda"
         Height          =   255
         Index           =   4
         Left            =   -73080
         TabIndex        =   26
         Top             =   3960
         Width           =   2655
      End
      Begin VB.OptionButton Opnalim 
         Caption         =   "Importados"
         Height          =   195
         Index           =   2
         Left            =   -71400
         TabIndex        =   25
         Top             =   2160
         Width           =   2895
      End
      Begin VB.OptionButton Opnalim 
         Caption         =   "Nacionales"
         Height          =   195
         Index           =   1
         Left            =   -71400
         TabIndex        =   24
         Top             =   1560
         Width           =   2895
      End
      Begin VB.OptionButton Opnalim 
         Caption         =   "General"
         Height          =   195
         Index           =   0
         Left            =   -71400
         TabIndex        =   23
         Top             =   1080
         Width           =   2895
      End
      Begin VB.OptionButton Opnalim 
         Caption         =   "Inactivos"
         Height          =   195
         Index           =   3
         Left            =   -71400
         TabIndex        =   22
         Top             =   2640
         Width           =   2895
      End
      Begin VB.OptionButton Opnalim 
         Caption         =   "Por proveedor"
         Height          =   195
         Index           =   4
         Left            =   -71400
         TabIndex        =   21
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton CmdDifPro 
         Caption         =   "Ver Precios"
         Height          =   375
         Left            =   -67200
         TabIndex        =   20
         Top             =   5040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame fraTrasl 
         Caption         =   "Traslados "
         ForeColor       =   &H80000002&
         Height          =   4095
         Left            =   -74400
         TabIndex        =   1
         Top             =   1185
         Width           =   9255
         Begin VB.OptionButton OptTrasl 
            Caption         =   "Salida de productos en traslados"
            Height          =   375
            Index           =   4
            Left            =   1320
            TabIndex        =   77
            Top             =   3000
            Width           =   2055
         End
         Begin VB.OptionButton OptTrasl 
            Caption         =   "Traslados por gondolero"
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   15
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CheckBox chkVolumen 
            Caption         =   "&Volúmen"
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
            Height          =   255
            Left            =   6480
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton OptTrasl 
            Caption         =   "Traslados pendientes de enviarse"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   6
            Top             =   1800
            Width           =   2895
         End
         Begin VB.OptionButton OptTrasl 
            Caption         =   "Productos sin desplazamiento en envios"
            Height          =   615
            Index           =   1
            Left            =   1320
            TabIndex        =   3
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chkPapeleria 
            Caption         =   "&Papeleria"
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
            Left            =   6480
            TabIndex        =   4
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton OptTrasl 
            Caption         =   "Envios a tiendas y franquicias"
            Height          =   495
            Index           =   0
            Left            =   1320
            TabIndex        =   2
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
      End
      Begin VB.PictureBox Rpt 
         Height          =   480
         Left            =   -74760
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   128
         Top             =   1365
         Width           =   1200
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Proveedor"
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
         Index           =   6
         Left            =   -72840
         TabIndex        =   42
         Top             =   1200
         Width           =   6375
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Producto"
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
         Index           =   7
         Left            =   -72840
         TabIndex        =   41
         Top             =   4680
         Width           =   6255
      End
   End
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   8010
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "                                                                                   Alt + tecla resaltada activa opción"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cArch As String
Private ntext As Integer
Private rsttemp As ADODB.Recordset

Private Sub sinmvto()
Dim CNP As ADODB.Connection
Me.stb1.SimpleText = "Espere, Generando Reporte de Productos "
  Set CNP = New Connection
  CNP.ConnectionString = cCadConex
  CNP.ConnectionTimeout = 0
  CNP.CommandTimeout = 0
  CNP.Open
CNP.Execute "DELETE FROM MVTOS"
If MsgBox("DESEAS VER EL REPORTE BASADO EN EL ACUMULADO DE VENTAS", vbInformation + vbYesNo) = vbYes Then
   cejecuta = "mvtoprod '" & dtfecini.Value & "','" & dtfecfin.Value & "',0"
   cencab = "PRODUCTOS SIN DESPLAZAMIENTO BASADO EN ACUMULADO DE VENTAS DEL " & Me.dtfecini.Value & " AL " & Me.dtfecfin.Value
   CNP.Execute "INSERT INTO MVTOS SELECT DESCRIPC,CONTENID,MEDIDA,PAQUETES,INCANT,INCANTPZA,INPROD,CLAPROVE " & _
                "FROM INVENTARIO,TFPRODUC " & _
                "WHERE CONSEC = INPROD AND ( INCANT > 0 OR INCANTPZA > 0 )  AND not INPROD  IN " & _
                "( SELECT MAX(consec) FROM VENGRAL " & _
                "WHERE  fecha between  '" & Me.dtfecini.Value & "' and  '" & DateAdd("d", 1, dtfecfin.Value) & "' " & _
                "group by consec)"

Else
   cejecuta = "mvtoprod '" & dtfecini.Value & "','" & dtfecfin.Value & "',1"
   cencab = "PRODUCTOS SIN DESPLAZAMIENTO BASADO EN SALIDAS DE BODEGA DEL " & Me.dtfecini.Value & " AL " & Me.dtfecfin.Value
   CNP.Execute "INSERT INTO MVTOS SELECT DESCRIPC,CONTENID,MEDIDA,PAQUETES,INCANT,INCANTPZA,INPROD,CLAPROVE " & _
                "From INVENTARIO, TFPRODUC " & _
                "WHERE CONSEC = INPROD AND ( INCANT > 0 OR INCANTPZA > 0 )  AND not INPROD  IN " & _
                    "( SELECT MAX(DT_PRODUCTO) " & _
                    "From DETALLETRASLADO, TRASLADOS " & _
                    "Where DT_CLAVE = T_CLAVE And t_enviado = 1 And t_entrada = 0 And t_motivocancela Is Null " & _
                    "and T_fecha between  '" & Me.dtfecini.Value & "'" & " and  '" & Me.dtfecfin.Value & "' " & _
                    "group by dt_producto)"

End If

'CNP.Execute (cejecuta)
fecha = dtfecini.Value
fecha1 = dtfecfin.Value
cARcRpt = "\mvtoprod.rpt"
frmAreaRecibo.cr1.Connect = cn
frmAreaRecibo.cr1.ReportFileName = App.Path & cARcRpt
frmAreaRecibo.cr1.WindowTitle = "Reporte de Productos Sin desplazamientos"
frmAreaRecibo.cr1.Formulas(0) = "ENCAB = '" & cencab & "'"
frmAreaRecibo.cr1.Action = 1
CNP.Close
Set CNP = Nothing
Exit Sub
cencab = "Periodo del  " & dtfecini.Value & " AL " & dtfecfin.Value
Open App.Path & "\MVTOPROD.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
Print #1, "RELACION DE PRODUCTOS SIN MOVIMIENTO "; UCase(Format(date, "long date"))
Print #1, cencab
Print #1, "=========================================================================================="
Print #1, "CAJAS    PIEZAS        PRODUCTO "
Print #1, "=========================================================================================="
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
prove = " "
rs.Open "SELECT descripC,contenid,medida,paquetes,incant,incantpza,claprove,nomprove  FROM catprov, MVTOS where prove = claprove order by claprove ", cn, adOpenKeyset, adLockOptimistic, adCmdText
While Not rs.EOF
   If prove <> rs!claprove Then
      Print #1, " "
      Print #1, " "
      Print #1, "  " & rs!claprove & vbTab & vbTab & vbTab & rs!NOMPROVE
      Print #1, " "
      prove = rs!claprove
   End If
   Print #1, "  "; rs!InCant & vbTab & rs!InCantPza & vbTab & rs!descripc & " " & rs!PAQUETES & " x " & rs!CONTENID & " " & rs!medida
   rs.MoveNext
Wend
Close #1
Set rs = Nothing
Handle = Shell("NOTEPAD " & App.Path & "\MVTOPROD.TXT", 1)
Exit Sub
End Sub


Private Sub ordenaventas()
' Macro1 Macro
' Macro grabada el 30/04/2001 por Moises Leon
    Cells.Select
    Cells.EntireColumn.AutoFit
    Selection.NumberFormat = "m/d/yyyy"
    Columns("E:E").Select
    Selection.NumberFormat = "#,##0.000"
    Range("D10").Select
    'ActiveWorkbook.Save
    Columns("H:J").Select
    Selection.NumberFormat = "#,##0.00"
    Columns("G:G").Select
    Selection.NumberFormat = "0.00"
    Cells.Select
    Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(7, 10), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    Range("B11").Select
End Sub
Private Sub btnreport_Click()
Select Case TabRep.Tab
    Case 0
        Call CredMayoreo
    Case 2
        Call Pedprov
    Case 3
        Call RPTEXIST
    Case 4
        Call reptraslados
    Case 5
        Call rptVentas
    Case 6
        Call repprod
    Case 7
        Call catprov
    Case 8
        Call rptagentes
End Select
End Sub

Private Sub rptagentes()
Dim cConRpt As String
'On Error GoTo Error:
Rpt.Formulas(1) = ""
Rpt.DataFiles(0) = ""
cMensaje = stb1.SimpleText
Cadconrpt = cCadConex
If OptAgente(0).Value = True Then 'Todas las preventas facturadas y no facturadas
   If Trim(cmbAgentes.Text) <> "" Then cCond = " AND VENTAS.agente = " & Mid(cmbAgentes.Text, InStr(1, cmbAgentes.Text, "|") + 1)
   crpt = "\Agtevta.rpt"
   cEnca = "PREVENTAS REGISTRADAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
   cadsql = "SELECT VENTAS.fecha, VENTAS.agente, VENTAS.credito, VENTAS.folpreventa, " & _
                    "VENTAS_DET.importe, VENTAS_DET.cancelado, " & _
                    "CATCLIENTE.cnombre " & _
            "FROM PITICO.Dbo.VENTAS VENTAS, " & _
                  "PITICO.dbo.VENTAS_DET VENTAS_DET, " & _
                  "PITICO.dbo.CATCLIENTE CATCLIENTE " & _
            "WHERE VENTAS.noventa = VENTAS_DET.noventa AND " & _
                   "VENTAS.agente = CATCLIENTE.cclave AND " & _
                   "VENTAS.fecha >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND " & _
                   "VENTAS.fecha <= '" & Format(dtfecfin.Value + 1, "yyyy-dd-mm") & "' AND " & _
                   "VENTAS.credito = 1 AND VENTAS_DET.cancelado = 0 " & cCond
   Rpt.SQLQuery = cadsql
   'MsgBox cadsql
ElseIf OptAgente(1).Value = True Then           'Ventas facturadas a Agentes
    crpt = "\AgteFac.rpt"
    cEnca = "FACTURADAS GENERADAS A CREDITO DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    cresp = MsgBox("DESEAS VER EL REPORTE EN RESUMEN?", vbYesNo + vbDefaultButton1 + vbQuestion)
    If cresp = vbYes Then
       Rpt.SectionFormat(0) = "DETAIL;F;F;F;X;X;X;X"
       Rpt.SectionFormat(1) = "GH1;T;X;X;X;X;X;X"
       Rpt.SectionFormat(2) = "GF1;T;X;X;X;X;X;X"
       Rpt.SectionFormat(3) = "GH2;T;X;X;X;X;X;X"
       Rpt.SectionFormat(4) = "GF2;F;X;X;X;X;X;X"
       Rpt.SectionFormat(5) = "GH3;F;X;X;X;X;X;X"
       Rpt.SectionFormat(6) = "GF3;F;X;X;X;X;X;X"
    Else
       Rpt.SectionFormat(0) = "DETAIL;F;F;F;X;X;X;X"
       Rpt.SectionFormat(1) = "GH1;T;X;X;X;X;X;X"
       Rpt.SectionFormat(2) = "GF1;T;X;X;X;X;X;X"
       Rpt.SectionFormat(3) = "GH2;T;X;X;X;X;X;X"
       Rpt.SectionFormat(4) = "GF2;T;X;X;X;X;X;X"
       Rpt.SectionFormat(5) = "GH3;T;X;X;X;X;X;X"
       Rpt.SectionFormat(6) = "GF3;F;X;X;X;X;X;X"
    End If
    If Trim(cmbAgentes.Text) <> "" Then
       condrpt = " AND VENTAS.agente = " & Mid(cmbAgentes.Text, InStr(1, cmbAgentes.Text, "|") + 1)
    Else
       condrpt = ""
    End If
    cadsql = "SELECT VENTAS.fecha, VENTAS.agente, VENTAS.folpreventa," & _
                     "FACVENTA.facfecha, FACVENTA.iva, FACVENTA.ieps, FACVENTA.cobrado, " & _
                     "FACVENTA_DET.importe, FACVENTA_DET.factura, FACVENTA_DET.serie, FACVENTA_DET.fecha_det " & _
             "FROM pitico.dbo.VENTAS VENTAS, " & _
                   "pitico.dbo.FACVENTA FACVENTA, " & _
                   "PITICO.dbo.FACVENTA_DET FACVENTA_DET " & _
             "WHERE VENTAS.noventa = FACVENTA.noventa AND " & _
                   "FACVENTA.numfactura = FACVENTA_DET.factura AND " & _
                   "FACVENTA.serie = FACVENTA_DET.serie AND " & _
                   "FACVENTA_DET.importe > 0. AND FACVENTA_DET.serie = 'B' AND " & _
                   "FACVENTA_DET.fecha_det >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND " & _
                   "FACVENTA_DET.fecha_det <= '" & Format(dtfecfin.Value, "yyyy-dd-mm") & "'" & condrpt
    Rpt.SQLQuery = cadsql
ElseIf OptAgente(2).Value = True Then    'Utilidad por agente
   If Trim(cmbAgentes.Text) <> "" Then cCond = " AND VENTAS.agente = " & Mid(cmbAgentes.Text, InStr(1, cmbAgentes.Text, "|") + 1)
    crpt = "\Agtediau.rpt"
    cEnca = "VENTAS A TRAVES DE AGENTE DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    Rpt.SQLQuery = "SELECT VENTAS.fecha, VENTAS.agente, VENTAS.credito, VENTAS.folpreventa, " & _
                     "VENTAS_DET.importe, VENTAS_DET.cancelado, VENTAS_DET.prebajo, " & _
                     "CATCLIENTE.cnombre " & _
             "FROM pitico.dbo.VENTAS VENTAS, " & _
                    "pitico.dbo.VENTAS_DET VENTAS_DET, " & _
                    "pitico.dbo.CATCLIENTE CATCLIENTE " & _
             "WHERE VENTAS.noventa = VENTAS_DET.noventa AND " & _
                    "VENTAS.agente = CATCLIENTE.cclave AND " & _
                    "VENTAS.fecha >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND VENTAS.fecha <='" & Format(dtfecfin.Value, "yyyy-dd-mm") & "' AND " & _
                    "VENTAS.agente <> '' AND ventas.folpreventa > 0 AND " & _
                    "VENTAS_DET.cancelado = 0 " & cCond
ElseIf OptAgente(5).Value = True Then
    If Trim(cmbAgentes.Text) <> "" Then cCond = " AND VENTAS.agente = " & Mid(cmbAgentes.Text, InStr(1, cmbAgentes.Text, "|") + 1)
    cEnca = "VENTAS DETALLADAS POR AGENTE DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    crpt = "\Agtipvta.rpt"
    Rpt.SQLQuery = "SELECT VENTAS.fecha, VENTAS.agente, VENTAS.credito, " & _
                            "VENTAS_DET.importe, VENTAS_DET.cancelado, VENTAS_DET.prebajo, " & _
                            "CATCLIENTE.cnombre,TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & _
                   "FROM pitico.dbo.VENTAS VENTAS, " & _
                        "pitico.dbo.VENTAS_DET VENTAS_DET, " & _
                        "pitico.dbo.CATCLIENTE CATCLIENTE, " & _
                        "PITICO.dbo.TFPRODUC TFPRODUC " & _
                   "WHERE VENTAS.noventa = VENTAS_DET.noventa AND " & _
                        "VENTAS.agente = CATCLIENTE.cclave AND " & _
                        "VENTAS.fecha >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND VENTAS.fecha <='" & Format(dtfecfin.Value, "yyyy-dd-mm") & "' AND " & _
                        "VENTAS.agente <> '' AND VENTAS.folpreventa > 0 AND " & _
                        "VENTAS_DET.cancelado = 0 AND " & _
                        "VENTAS_DET.cl_producto = TFPRODUC.CONSEC" & cCond
ElseIf OptAgente(3).Value = True Or OptAgente(4).Value = True Then
    CTEAGTE = IIf(OptAgente(3).Value = True, cmbCliente.Text, cmbAgentes.Text)
    condcte = " CATCLIENTE.ctipo = " & IIf(OptAgente(3).Value = True, 0, 1) & " AND "
    If Trim(CTEAGTE) <> "" Then cCond = " FACVENTA.faccliente = " & Mid(CTEAGTE, InStr(1, CTEAGTE, "|") + 1) & " AND "
    If MsgBox("DESEAS VER SOLAMENTE FACTURAS PENDIENTES DE COBRO?", vbQuestion + vbYesNo, "Facturas") = vbYes Then cCond = cCond + " FACVENTA.cobrado = 0 and "
    cEnca = "CREDITO DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    crpt = "\agtecred.rpt"
    Rpt.SQLQuery = "SELECT FACVENTA.facfecha, FACVENTA.total, FACVENTA.iva, FACVENTA.ieps, FACVENTA.cancelado, FACVENTA.numfactura, FACVENTA.serie, FACVENTA.cobrado, FACVENTA.rfc, FACVENTA.TotAbono, FACVENTA.porpagar, " & _
                           "CATCLIENTE.cclave, CATCLIENTE.cnombre, CATCLIENTE.cTipo  " & _
                   "FROM PITICO.dbo.FACVENTA FACVENTA, " & _
                         "PITICO.dbo.CATCLIENTE CATCLIENTE " & _
                   "WHERE FACVENTA.faccliente = CATCLIENTE.cclave AND " & _
                         "FACVENTA.cancelado = 0  AND " & cCond & condcte & _
                         "FACVENTA.facfecha >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND FACVENTA.facfecha <='" & Format(dtfecfin.Value, "yyyy-dd-mm") & "'"
ElseIf OptAgente(6).Value = True Then
    If Trim(cmbCliente.Text) <> "" Then cCond = " FACVENTA.faccliente = " & Mid(cmbCliente.Text, InStr(1, cmbCliente.Text, "|") + 1) & " AND "
    cEnca = "FACTURAS GENERADAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    crpt = "\agtecred1.rpt"
    Rpt.SQLQuery = "SELECT FACVENTA.facfecha, FACVENTA.total, FACVENTA.iva, FACVENTA.ieps, FACVENTA.cancelado, FACVENTA.numfactura, FACVENTA.serie, FACVENTA.rfc, FACVENTA.faccobro, FACVENTA.TotAbono, FACVENTA.porpagar, " & _
                           "CATCLIENTE.cclave, CATCLIENTE.cnombre, CATCLIENTE.ctipo, " & _
                           "FACVENTA_DET.Cantidad, FACVENTA_DET.cantidadp, FACVENTA_DET.costo, FACVENTA_DET.costop, FACVENTA_DET.importe, TFPRODUC.DESCRIPC " & Chr(13) & _
                   "FROM PITICO.dbo.FACVENTA FACVENTA, " & _
                        "PITICO.dbo.CATCLIENTE CATCLIENTE, " & _
                        "pitico.dbo.FACVENTA_DET FACVENTA_DET, " & _
                        "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                   "WHERE FACVENTA.faccliente = CATCLIENTE.cclave AND FACVENTA.numfactura = FACVENTA_DET.factura AND " & cCond & _
                        "FACVENTA.serie = FACVENTA_DET.serie AND FACVENTA_DET.Producto = TFPRODUC.CONSEC AND " & _
                        "FACVENTA.facfecha >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND FACVENTA.facfecha <='" & Format(dtfecfin.Value, "yyyy-dd-mm") & "' AND FACVENTA.cancelado = 0"
ElseIf OptAgente(7).Value = True Then
    cEnca = "UTILIDAD POR AGENTE DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    crpt = "\agteuti.rpt"
    Rpt.SQLQuery = "SELECT VENTAS.agente, " & _
                            "CATCLIENTE.cnombre, " & _
                            "FACVENTA.facfecha, FACVENTA.serie, " & _
                            "FACVENTA_DET.Cantidad, FACVENTA_DET.cantidadp, FACVENTA_DET.costo, FACVENTA_DET.costop, FACVENTA_DET.importe " & Chr(13) & _
                     "FROM pitico.dbo.VENTAS VENTAS, " & _
                            "PITICO.dbo.CATCLIENTE CATCLIENTE, " & _
                            "PITICO.dbo.FACVENTA FACVENTA, " & _
                            "PITICO.dbo.FACVENTA_DET FACVENTA_DET " & Chr(13) & _
                     "WHERE VENTAS.agente = CATCLIENTE.cclave AND " & _
                            "VENTAS.noventa = FACVENTA.noventa AND " & _
                            "FACVENTA.numfactura = FACVENTA_DET.factura AND " & _
                            "FACVENTA.serie = FACVENTA_DET.serie AND " & _
                            "FACVENTA_det.fecha_det >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND FACVENTA_det.FECHA_det <= '" & Format(dtfecfin.Value, "yyyy-dd-mm") & "'" & _
                            ""
ElseIf OptAgente(8).Value = True Then
    stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
    stb1.Refresh
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Execute "DELETE FROM comparativo"
    'Preventas Base
    cn.Execute "INSERT INTO comparativo(agente,preventa,especial,NORMAL) SELECT agente, MAX(V.FOLPREVENTA),  SUM( CASE PREBAJO WHEN 1 THEN IMPORTE END ) PREBAJO, SUM( CASE PREBAJO WHEN 0 THEN IMPORTE END ) PRENORMAL " & _
            "FROM VENTAS_DET D, VENTAS V WHERE V.NOVENTA = D.NOVENTA AND FECHA >= '" & dtfecfin.Value & "' AND FECHA <= '" & dtfecfin.Value + 1 & "' AND D.CANCELADO = 0 AND CREDITO = 1 GROUP BY AGENTE"
    'Preventa Semana anterior
    rs.Open "SELECT agente, MAX(V.FOLPREVENTA) preventa,  SUM( CASE PREBAJO WHEN 1 THEN IMPORTE END ) PREBAJO, SUM( CASE PREBAJO WHEN 0 THEN IMPORTE END ) PRENORMAL " & _
            "FROM VENTAS_DET D, VENTAS V WHERE V.NOVENTA = D.NOVENTA AND FECHA >= '" & dtfecfin.Value - 7 & "' AND FECHA <= '" & dtfecfin.Value - 7 + 1 & "' AND D.CANCELADO = 0 AND CREDITO = 1 GROUP BY AGENTE", cn, adOpenForwardOnly, adLockOptimistic, admcdtext
    While Not rs.EOF
        cn.Execute "UPDATE comparativo SET normala  = " & IIf(IsNull(rs!prenormal), 0, rs!prenormal) & ", especiala = " & IIf(IsNull(rs!Prebajo), 0, rs!Prebajo) & " WHERE agente = " & rs!agente
        rs.MoveNext
    Wend
    rs.Close
    cadsql = "SELECT SUM( CASE prebajo WHEN 1 THEN importe END ) PREBAJO, SUM( CASE prebajo WHEN 0 THEN importe END ) PRENORMAL, " & _
                    "MAX(V.folpreventa) PREVENTA , AGENTE  FROM VENTAS_DET D, VENTAS V WHERE V.NOVENTA = D.NOVENTA AND FECHA >= '" & dtfecini.Value & "' AND FECHA <= '" & dtfecfin.Value + 1 & "' AND D.CANCELADO = 0 AND ("
    fecha = dtfecfin.Value
    N = 0
    While fecha >= dtfecini.Value
        cadsql = cadsql + " DAY(FECHA) = " & Day(fecha) & " OR"
        N = N + 1
        fecha = fecha - 7
    Wend
    cadsql = Mid(cadsql, 1, Len(cadsql) - 3) & ") AND CREDITO = 1 GROUP BY AGENTE"
    'MsgBox CADSQL
    rs.Open cadsql, cn, adOpenStatic, adLockOptimistic, admcdtext
    While Not rs.EOF
        cn.Execute "UPDATE comparativo SET normalp  = " & IIf(IsNull(rs!prenormal), 0, rs!prenormal / N * 1.1) & ", especialp = " & IIf(IsNull(rs!Prebajo), 0, rs!Prebajo / N * 1.1) & " WHERE agente = " & rs!agente
        rs.MoveNext
    Wend
    cn.Execute "UPDATE comparativo SET normal = 0 where normal is null"
    cn.Execute "UPDATE comparativo SET especial = 0 where especial is null "
    cn.Execute "UPDATE comparativo SET normala = 0 where normala is null "
    cn.Execute "UPDATE comparativo SET especiala = 0 where especiala is null "
    crpt = "\agtecomp.rpt"
    cEnca = "COMPARATIVO DE VENTAS POR DIA Y AGENTE DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    Rpt.Formulas(1) = "NOTA = '*** EL COMPARATIVO SE ESTA OBTENIENDO DE " & N & " " & UCase(Format(Me.dtfecfin.Value, "dddd")) & " ANTERIORES'"
ElseIf OptAgente(9).Value = True Then
    If MsgBox("VER SOLAMENTE PENDIENTES DE LIQUIDAR", vbQuestion + vbYesNo, "Preventas") = vbYes Then
       COND = "PREVENTAS.impliq = 0 AND PREVENTAS.impsurt > 0 AND "
    Else
       COND = ""
    End If
    Rpt.SQLQuery = "SELECT VENTAS.agente, VENTAS.folpreventa, " & _
                    "CATCLIENTE.cnombre, PREVENTAS.folio, PREVENTAS.impcapt, PREVENTAS.impsurt, PREVENTAS.impliq, PREVENTAS.fecha " & _
             "FROM PITICO.dbo.VENTAS VENTAS, " & _
                   "PITICO.dbo.CATCLIENTE CATCLIENTE, " & _
                   "PITICO.dbo.PREVENTAS PREVENTAS " & _
             "WHERE  VENTAS.agente = CATCLIENTE.cclave AND " & _
                   "VENTAS.folpreventa = PREVENTAS.folio AND " & COND & _
                   "PREVENTAS.fecha >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND PREVENTAS.fecha <='" & Format(dtfecfin.Value, "yyyy-dd-mm") & "' " & Chr(13) & _
             "ORDER BY VENTAS.agente ASC, VENTAS.folpreventa ASC"
     crpt = "\agtepvta.rpt"
     cEnca = "PREVENTAS REGISTRADAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
ElseIf OptAgente(10).Value = True Then
    If cmbAgentes.Text = "" Then
       MsgBox "Es necesario especificar el agente", vbInformation, "Agentes"
       Exit Sub
    End If
    cEnca = "RUTA DEL AGENTE " & cmbAgentes.Text
    crpt = "\Ruta.rpt"
    Rpt.DataFiles(0) = "P:\PREVENTA\" & Trim(Mid(cmbAgentes.Text, InStr(1, cmbAgentes.Text, "|") + 1)) & ".mdb"
End If
'MsgBox Rpt.SQLQuery
stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
stb1.Refresh
Rpt.Connect = cCadConex
Rpt.ReportFileName = App.Path & crpt
Rpt.WindowTitle = "Reporte de ventas por agente"
Rpt.Formulas(0) = "ENCAB = '" & cEnca & "'"
Rpt.Action = 1
stb1.SimpleText = cMensaje
stb1.Refresh
cCadConex = Cadconrpt
Exit Sub
Error:
  MsgBox Err.Description, vbCritical
  Unload Me
End Sub
Private Sub rptprodina()
Adoexcel.CommandType = adCmdText
    Adoexcel.CursorType = adOpenKeyset
    Adoexcel.ConnectionString = cCadConex
    cadenaconexion = "SELECT descripc,contenid,medida,paquetes,precio1,precosto, claprove,nomprove,incant,incantpza,barraspza,consec FROM TFPRODUC,PREPROD,INVENTARIO, CATPROV WHERE  CONSEC = PRECLAVE AND tfproduc.ACTIVO = 0  AND CONSEC = INPROD  and  CLAPROVE = PROVE AND FECACT BETWEEN '" & dtfecini.Value & "'  AND '" & dtfecfin.Value & "' ORDER BY DESCRIPC "
    'MsgBox cadenaconexion
    Adoexcel.RecordSource = cadenaconexion
    Adoexcel.Refresh
    If Adoexcel.Recordset.BOF And Adoexcel.Recordset.EOF Then
        MsgBox "No existen Productos con la condición especificada ... ", vbInformation
        Exit Sub
    Else
        Call SELECTFILE
        'CONFIGURACION DEL ARCHIVO DE EXCEL
        Dim Targetdir As String
        Contador_Project_Planning = 1
        Contador_Project_Selected = 1
        Contador_Project_Files = 1
        MousePointer = vbHourglass
        Targetdir = cArch
        Set Appxl = CreateObject("Excel.Application")
        Appxl.Visible = True
        Appxl.Workbooks.Add
        ProBar1.Visible = True
        ProBar1.Min = 0
        ProBar1.Max = Adoexcel.Recordset.RecordCount
        contador_Fila = 1
        ' Renombro las hojas de Excel y le pongo los encabezados para introducir los datos
        Appxl.Sheets("hoja1").Select
        Appxl.Sheets("hoja1").Name = "INACTIVOS"
        Appxl.Cells(Contador_Project_Selected, 1).Value = "PROV"
        Appxl.Cells(Contador_Project_Selected, 2).Value = "PROVEEDOR"
        Appxl.Cells(Contador_Project_Selected, 3).Value = "PRODUCTO"
        Appxl.Cells(Contador_Project_Selected, 4).Value = "PAQUETES"
        Appxl.Cells(Contador_Project_Selected, 5).Value = "CONT."
        Appxl.Cells(Contador_Project_Selected, 6).Value = "MEDIDA"
        Appxl.Cells(Contador_Project_Selected, 7).Value = "CAJAS"
        Appxl.Cells(Contador_Project_Selected, 8).Value = "PIEZAS"
        Appxl.Cells(Contador_Project_Selected, 9).Value = "PRE.COSTO"
        Appxl.Cells(Contador_Project_Selected, 10).Value = "PRE. AUT."
        Appxl.Cells(Contador_Project_Selected, 11).Value = "BARRASPZA"
        Appxl.Cells(Contador_Project_Selected, 12).Value = "CLAVE"
        Contador_Project_Selected = Contador_Project_Selected + 1
        Appxl.Sheets("INACTIVOS").Select
        While Not Adoexcel.Recordset.EOF
            Appxl.Cells(Contador_Project_Selected, 1).Value = Adoexcel.Recordset!claprove
            Appxl.Cells(Contador_Project_Selected, 2).Value = Adoexcel.Recordset!NOMPROVE
            Appxl.Cells(Contador_Project_Selected, 3).Value = Adoexcel.Recordset!descripc
            Appxl.Cells(Contador_Project_Selected, 4).Value = Adoexcel.Recordset!PAQUETES
            Appxl.Cells(Contador_Project_Selected, 5).Value = Adoexcel.Recordset!CONTENID
            Appxl.Cells(Contador_Project_Selected, 6).Value = Adoexcel.Recordset!medida
            Appxl.Cells(Contador_Project_Selected, 7).Value = Adoexcel.Recordset!InCant
            Appxl.Cells(Contador_Project_Selected, 8).Value = Adoexcel.Recordset!InCantPza
            Appxl.Cells(Contador_Project_Selected, 9).Value = Adoexcel.Recordset!PRECOSTO
            Appxl.Cells(Contador_Project_Selected, 10).Value = Adoexcel.Recordset!precio1
            Appxl.Cells(Contador_Project_Selected, 11).Value = Adoexcel.Recordset!barraspza
            Appxl.Cells(Contador_Project_Selected, 12).Value = Adoexcel.Recordset!CONSEC
            Contador_Project_Selected = Contador_Project_Selected + 1
            v = v + 1
            ProBar1.Value = v
            Adoexcel.Recordset.MoveNext
        Wend
        Adoexcel.Recordset.Close
    End If
    Appxl.ActiveWorkbook.SaveAs Targetdir
    Set Appxl = Nothing
    MousePointer = vbDefault
    MsgBox "PROCESO FINALIZADO...", vbInformation
    ProBar1.Visible = False
    Exit Sub
End Sub

Private Sub CredMayoreo()
On Error GoTo Error:
cMensaje = stb1.SimpleText
If Optcred(0).Value = True Then
    crpt = App.Path & "\ctecred.rpt"
    If MsgBox("DESEAS COSULTAR CLIENTES SOLAMENTE CON CREDITO", vbYesNo + vbQuestion, "Tipo") = vbYes Then
       Rpt.Formulas(1) = "FORMSELEC = {CATCLIENTE.Ccredito}  = 1"
       cEnca = "CARTERA DE CLIENTES CON DERECHO A CREDITO"
    Else
       Rpt.Formulas(1) = "FORMSELEC = {CATCLIENTE.Ccredito}  = 0"
       cEnca = "CARTERA DE CLIENTES"
    End If
ElseIf Optcred(1).Value = True Then
    cEnca = "CREDITOS EN BODEGA PUERTO ESCONDIDO"
    crpt = "P:\buzon\credito55.rpt"
ElseIf Optcred(2).Value = True Then
    cEnca = "CARTERA DE CLIENTES CON DERECHO A CREDITO"
    crpt = "P:\buzon\ctecred55.rpt"
End If
stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
stb1.Refresh
Rpt.Connect = cCadConex
Rpt.ReportFileName = crpt
Rpt.WindowTitle = "Reporte de ventas por agente"
Rpt.Formulas(0) = "ENCAB = '" & cEnca & "'"
Rpt.Action = 1
stb1.SimpleText = cMensaje
stb1.Refresh
Exit Sub
Error:
  MsgBox Err.Description, vbCritical
  Unload Me
End Sub


Private Sub repprod()
On Error GoTo Error:
crpt = "\NalImpor.rpt"
If Opnalim(0).Value = True Then
    condrpt = "{tfproduc.CONSEC} <> '00000' "
    Enca = "LISTADO DE PRODUCTOS NACIONALES E IMPORTADOS"
ElseIf Opnalim(1).Value = True Then
    condrpt = "{tfproduc.PROCEDENCIA} = 1"
    Enca = "LISTADO DE PRODUCTOS NACIONALES"
ElseIf Opnalim(2).Value = True Then
    condrpt = "{tfproduc.PROCEDENCIA} = 0"
    Enca = "LISTADO DE PRODUCTOS DE IMPORTACION"
ElseIf Opnalim(3).Value = True Then
    condrpt = "{TFPRODUC.activo} = 0"
    Enca = "LISTADO DE PRODUCTOS INACTIVOS POR PROVEEDOR"
    crpt = "\PROINAC.rpt"
ElseIf Opnalim(4).Value = True Then
    Enca = "LISTADO DE PRODUCTOS POR PROVEEDOR"
    crpt = "\proxprov.rpt"
ElseIf Opnalim(5).Value = True Then
    Call rptprodina
    Exit Sub
    Enca = "LISTADO DE PRODUCTOS INACTIVADOS "
    crpt = "\prodina.rpt"
ElseIf Opnalim(6).Value = True Then
    Enca = "LISTADO DE PRECIOS POR PROVEEDOR"
    crpt = "\ListProv.rpt"
End If
Rpt.Connect = cCadConex
Rpt.ReportFileName = App.Path & crpt
If Opnalim(4).Value Or Opnalim(6).Value Then
     Rpt.WindowTitle = "Productos del proveedor " & cmbProvPro.Text
     N = Len(cmbProvPro.Text) - 4
     cveprov = Trim(Mid(cmbProvPro.Text, N + 1))
     Rpt.Formulas(0) = "FORMSELEC = {TFPRODUC.CLAPROVE} = '" & cveprov & "'"
Else
     Rpt.Formulas(0) = "FORMSELEC = " & condrpt
     Rpt.Formulas(1) = "ENCAB = '" & Enca & "'"
     Rpt.WindowTitle = Enca
End If
Rpt.Action = 1
Exit Sub
Error:
   MsgBox Err.Description
   Unload Me
End Sub

Private Sub btnsal_Click(Index As Integer)
Unload Me
End Sub

Private Sub btnventas_Click(Index As Integer)
Select Case Index
Case 4
    Dim cntmp As ADODB.Connection
    Dim RST As ADODB.Recordset
    Set cntmp = New ADODB.Connection
    cntmp.ConnectionTimeout = 0
    cntmp.CommandTimeout = 0
    cntmp.ConnectionString = cCadConex
    cntmp.Open
    Set RST = New ADODB.Recordset
    cmensa = stb1.SimpleText
    stb1.SimpleText = Space(20) & "Buscando acumulados del período especificado"
    RST.Open "SELECT caja, SUM(precio) totcaja,FECHA FROM vengral WHERE FECHA >='" & Me.dtfecini.Value & "' AND FECHA <= '" & Me.dtfecfin.Value & "' GROUP BY CAJA,FECHA ORDER BY FECHA", cntmp, adOpenForwardOnly, adLockOptimistic, adCmdText
    Open App.Path & "\ACU" & Mid(Trim(Mid(cSucursal, 3)), 1, 3) & ".TXT" For Output As #1 ' Abre el archivo para operaciones de salida.
    Print #1, "PUNTOS DE VENTA DE LOS QUE EXISTE ACUMULADO DE VENTAS"
    Print #1, ""
    Print #1, " ESTACION   FECHA   IMPORTE VENTA"
    Print #1, ""
    If Not (RST.BOF And RST.EOF) Then FECANT = RST!fecha
    TOTDIA = 0: TotVta = 0
    While Not RST.EOF
        If FECANT = RST!fecha Then
           TOTDIA = TOTDIA + RST!TOTCAJA
        Else
           Print #1, " TOTAL DIA" & Space(3) & Format(TOTDIA, "$###,###,##0.00")
           Print #1, ""
           TotVta = TotVta + TOTDIA
           TOTDIA = 0
           TOTDIA = TOTDIA + RST!TOTCAJA
        End If
        Print #1, Space(3) + RST!Caja & Space(5) & RST!fecha & Space(5) & Format(RST!TOTCAJA, "$###,##0.00")
        FECANT = RST!fecha
        RST.MoveNext
    Wend
    If Not (RST.BOF And RST.EOF) Then
       Print #1, " TOTAL DIA" & Space(3) & Format(TOTDIA, "$###,###,##0.00")
       Print #1, ""
       Print #1, " TOTAL GENERAL" & Space(3) & Format(TotVta, "$###,###,##0.00")
    End If
    Close #1
    Handle = Shell("NOTEPAD " & App.Path & "\" & "\ACU" & Mid(Trim(Mid(cSucursal, 3)), 1, 3) & ".TXT", 1)
    RST.Close
    Set RST = Nothing
    cntmp.Close
    Set cntmp = Nothing
    stb1.SimpleText = cmensa
End Select
   
End Sub


Private Sub cmbagente_GotFocus()
Dim rsttemp As ADODB.Recordset
Set rsttemp = New ADODB.Recordset
If Me.cmbAgentes.ListCount = 0 Then
'cargo datos de la pestaña Agentes
 rsttemp.Open "SELECT * FROM Catcliente WHERE ctipo = 1", cCadConex, adOpenDynamic, adLockOptimistic, adCmdText
 While Not rsttemp.EOF
    If (Not IsNull(rsttemp!cNombre)) And Not (IsNull(rsttemp!cclave)) Then
        cmbAgentes.AddItem rsttemp!cNombre & "    " & rsttemp!cclave
    End If
    rsttemp.MoveNext
 Wend
 rsttemp.Close
 Set rsttemp = Nothing
 stb1.SimpleText = cMensaje
 stb1.Refresh
End If
End Sub

Private Sub cmbAgentes_GotFocus()
Dim rsttemp As ADODB.Recordset
Set rsttemp = New ADODB.Recordset
If Me.cmbAgentes.ListCount = 0 Then
'cargo datos de la pestaña Agentes
 rsttemp.Open "SELECT * FROM Catcliente WHERE ctipo = 1", cCadConex, adOpenDynamic, adLockOptimistic, adCmdText
 While Not rsttemp.EOF
    If (Not IsNull(rsttemp!cNombre)) And Not (IsNull(rsttemp!cclave)) Then
        cmbAgentes.AddItem rsttemp!cNombre & "    | " & rsttemp!cclave
    End If
    rsttemp.MoveNext
 Wend
 rsttemp.Close
 Set rsttemp = Nothing
 stb1.SimpleText = cMensaje
 stb1.Refresh
End If
End Sub

Private Sub cmbCliente_GotFocus()
If cmbCliente.ListCount = 0 Then
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open "SELECT distinct cclave,cnombre FROM facventa, catcliente WHERE cclave = faccliente and facfecha >= '" & dtfecini.Value & "' and facfecha <= '" & dtfecfin.Value & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
  While Not rs.EOF
     cmbCliente.AddItem rs!cNombre & "    |" & rs!cclave
     rs.MoveNext
  Wend
  rs.Close
  Set rs = Nothing
End If
End Sub


Private Sub cmdopcion_GotFocus(Index As Integer)
If Index = 0 Then Unload frmAreaRecibo
End Sub

Private Sub cmbprod_GotFocus()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If InStr(1, cmbProVta.Text, "|") > 0 Then
   rs.Open "SELECT CONSEC, Descripc, LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS PRESENT FROM tfproduc WHERE activo = 1 And claprove = '" & Trim(Mid(cmbProVta.Text, InStr(1, cmbProVta.Text, "|") + 1)) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   cmbprod.Clear
   While Not rs.EOF
      cmbprod.AddItem rs!descripc & "    " & rs!Present & "|   " & rs!CONSEC
      rs.MoveNext
   Wend
End If
Set rs = Nothing
End Sub

Private Sub RPTEXIST()
On Error GoTo Error:
'Valido datos
If OptExis(0).Value = True Then
   If txtExis.Text = "" Then
      MsgBox "NO PUEDE DEJAR EN BLANCO LA CANTIDAD", vbExclamation
      Exit Sub
   ElseIf Not IsNumeric(txtExis.Text) Then
      MsgBox "La cantidad debe ser numerica", vbExclamation
      Exit Sub
   End If
End If
For N = 0 To 3
    If cmbClas(N).Visible = True And cmbClas(N).Text = "" Then
        MsgBox "No puede dejar en blanco la condicion", vbExclamation
        Exit Sub
    End If
Next
'SE PONE EL INVENTARIO CON COSTOS
If PORCOSTO = True Then
  ccondrpt = "{INVENTARIO.inCant} > 0"
  Call INVCOSTO
  PORCOSTO = False
  Exit Sub
End If
'Inicia la construccion de encabezados y condicion del Rpt
crpt = "\Exist.rpt"
If OptExis(0).Value = True Then
   cEnca = "INVENTARIO CON EXISTENCIA " + cmbClas(0).Text + "  " + txtExis.Text
   ccondrpt = "{INVENTARIO.inCant} " + cmbClas(0).Text + "  " + txtExis.Text
ElseIf OptExis(1).Value = True Then
   cEnca = "INVENTARIO DE LA FAMILIA " & cmbClas(1).Text
   ccondrpt = " {FAMILIAS.fclave} = '" + Mid(cmbClas(1).Text, 1, 3) + "'"
ElseIf OptExis(2).Value = True Then
   cEnca = "INVENTARIO DE LA LINEA " & cmbClas(2).Text
   ccondrpt = "{LINEAS.sfclave} = '" + Trim(Mid(cmbClas(2).Text, Len(cmbClas(2).Text) - 5)) + "'"
   crpt = "\EXISTEN.rpt"
ElseIf OptExis(3).Value = True Then
   cEnca = "INVENTARIO DEL DEPARTAMENTO " & cmbClas(3).Text
   ccondrpt = "{DEPARTAMENTO.depclave} = '" + Mid(cmbClas(3).Text, 1, 3) + "'"
ElseIf OptExis(4).Value = True Then
   crpt = "\ExiPro.rpt"
   cEnca = "INVENTARIO DEL PROVEEDOR " & cmbProved.Text
   ccondrpt = "{TFPRODUC.claprove} = '" + Trim(Mid(cmbProved.Text, Len(cmbProved.Text) - 5)) + "'"
ElseIf OptExis(7).Value = True Then
   cmen = stb1.SimpleText
   cn.Execute "DELETE FROM exinoenv"
   For N = 0 To lstNoenv.ListCount - 1
       If lstNoenv.Selected(N) Then
          cn.Execute "INSERT INTO exinoenv SELECT inprod, incant, incantpza, tienda= '" & Trim(Mid(lstNoenv.List(N), 3, 20)) & "' FROM INVENTARIO WHERE incant > 0 AND NOT inprod IN (SELECT inprod FROM inventario" & Trim(Mid(lstNoenv.List(N), 1, 3)) & " WHERE incant > 0) "
          stb1.SimpleText = "Verificando existencias de sucursal " & lstNoenv.List(N)
         stb1.Refresh
       End If
   Next
   Rpt.WindowTitle = "Productos con existencia en Carbonera y que no tienen las tiendas en bodega"
   Rpt.ReportFileName = App.Path & "\exinoenv.rpt"
   Rpt.Action = 1
   stb1.SimpleText = cmen
   stb1.Refresh
   Exit Sub
ElseIf OptExis(6).Value = True Then
    Adoexcel.CommandType = adCmdText
    Adoexcel.CursorType = adOpenKeyset
    Adoexcel.ConnectionString = cCadConex
    If tipotienda = 1 Then
       cadenaconexion = "SELECT consec, sfdescrip AS linea, fdescrip AS familia,inprod, descripc, paquetes,STR(contenid) + ' ' + medida AS MEDIDA, incant, incantpza, precio1,precosto, Barraspza from inventario,preprod,tfproduc,lineas,familias " & _
                     "Where consec = inprod And inprod = preclave And (incant   " & cmbClas(0).Text & " " & Trim(txtExis.Text) & " Or incantpza > 0)  AND TFPRODUC.linea = lineas.sfclave AND sffamilia = fclave " & compInt & " ORDER BY descripc "
    Else
       cadenaconexion = "SELECT consec, sfdescrip AS linea, familia = '',inprod, descripc, paquetes,STR(contenid,10,3) + ' ' + medida AS MEDIDA, incant, incantpza, precio1,precosto, Barraspza from inventario,preprod,tfproduc,lineas " & _
                     "WHERE TFPRODUC.linea *= lineas.sfclave AND consec = inprod And inprod = preclave And (incant   " & cmbClas(0).Text & " " & Trim(txtExis.Text) & " Or incantpza > 0) " & compInt & " ORDER BY descripc "
    End If
    Adoexcel.RecordSource = cadenaconexion
    Adoexcel.Refresh
    
    If Adoexcel.Recordset.BOF And Adoexcel.Recordset.EOF Then
        MsgBox "No existen Productos con la condición especificada ... ", vbInformation
        Exit Sub
    Else
        Call SELECTFILE
        'CONFIGURACION DEL ARCHIVO DE EXCEL
        Dim Targetdir As String
        Contador_Project_Planning = 1
        Contador_Project_Selected = 1
        Contador_Project_Files = 1
        MousePointer = vbHourglass
        Targetdir = cArch
        Set Appxl = CreateObject("Excel.Application")  'run it
        Appxl.Visible = True
        Appxl.Workbooks.Add
         
        ProBar1.Visible = True
        ProBar1.Min = 0
        ProBar1.Max = Adoexcel.Recordset.RecordCount
        contador_Fila = 1
        ' Renombro las hojas de Excel y le pongo los encabezados para introducir los datos
        Appxl.Sheets("hoja1").Select
        Appxl.Sheets("hoja1").Name = "EXISTENCIAS"
        Appxl.Cells(Contador_Project_Selected, 1).Value = "LINEA"
        Appxl.Cells(Contador_Project_Selected, 2).Value = "FAMILIA"
        Appxl.Cells(Contador_Project_Selected, 3).Value = "PRODUCTO"
        Appxl.Cells(Contador_Project_Selected, 4).Value = "PAQUETES"
        Appxl.Cells(Contador_Project_Selected, 5).Value = "MEDIDA"
        Appxl.Cells(Contador_Project_Selected, 6).Value = "CAJAS"
        Appxl.Cells(Contador_Project_Selected, 7).Value = "PIEZAS"
        Appxl.Cells(Contador_Project_Selected, 8).Value = "PRE.COSTO"
        Appxl.Cells(Contador_Project_Selected, 9).Value = "PRE. AUT."
        Appxl.Cells(Contador_Project_Selected, 10).Value = "BARRASPZA"
        Appxl.Cells(Contador_Project_Selected, 11).Value = "CLAVE"
        Contador_Project_Selected = Contador_Project_Selected + 1
        Appxl.Sheets("EXISTENCIAS").Select
        While Not Adoexcel.Recordset.EOF
            Appxl.Cells(Contador_Project_Selected, 1).Value = Adoexcel.Recordset!linea
            Appxl.Cells(Contador_Project_Selected, 2).Value = Adoexcel.Recordset!familia
            Appxl.Cells(Contador_Project_Selected, 3).Value = Adoexcel.Recordset!descripc
            Appxl.Cells(Contador_Project_Selected, 4).Value = Adoexcel.Recordset!PAQUETES
            Appxl.Cells(Contador_Project_Selected, 5).Value = Adoexcel.Recordset!medida
            Appxl.Cells(Contador_Project_Selected, 6).Value = Adoexcel.Recordset!InCant
            Appxl.Cells(Contador_Project_Selected, 7).Value = Adoexcel.Recordset!InCantPza
            Appxl.Cells(Contador_Project_Selected, 8).Value = Adoexcel.Recordset!PRECOSTO
            Appxl.Cells(Contador_Project_Selected, 9).Value = Adoexcel.Recordset!precio1
            Appxl.Cells(Contador_Project_Selected, 10).Value = Adoexcel.Recordset!barraspza
            Appxl.Cells(Contador_Project_Selected, 11).Value = Adoexcel.Recordset!CONSEC
            Contador_Project_Selected = Contador_Project_Selected + 1
            v = v + 1
            ProBar1.Value = v
            Adoexcel.Recordset.MoveNext
        Wend
        Adoexcel.Recordset.Close
        Call ordenaexistencias
    End If
    'appXL.ActiveWorkbook.Save Targetdir
    Appxl.ActiveWorkbook.SaveAs Targetdir
    'Appxl.ActiveWorkbook.Close (False)
    'appxl.Application.Quit
    Set Appxl = Nothing
    MousePointer = vbDefault
    MsgBox "PROCESO FINALIZADO...", vbInformation
    ProBar1.Visible = False
    Exit Sub
ElseIf OptExis(8).Value = True Then
    Rpt.ReportFileName = App.Path & "\ventasft.rpt"
    Rpt.SQLQuery = "SELECT FACVENTA_DET.Cantidad, FACVENTA_DET.importe, FACVENTA_DET.fecha_det, FACVENTA_DET.rfc_det, " & _
                           "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC, INVENTARIO.incant " & _
                   "FROM pitico.dbo.FACVENTA_DET FACVENTA_DET, " & _
                       "pitico.dbo.TFPRODUC TFPRODUC, pitico.dbo.INVENTARIO INVENTARIO " & _
                   "WHERE FACVENTA_DET.Producto = TFPRODUC.CONSEC AND FACVENTA_DET.Producto = INVENTARIO.inprod AND " & _
                         "FACVENTA_DET.fecha_det >= '" & Format(DateAdd("M", -3, dtfecfin.Value), "yyyy-dd-mm") & "' AND FACVENTA_DET.fecha_det <= '" & Format(dtfecfin.Value, "yyyy-dd-mm") & "' AND FACVENTA_DET.rfc_det <> 'CANC999999999'"
    'MsgBox Rpt.SQLQuery
    Rpt.Formulas(0) = "ENCAB = 'VENTAS DEL " & DateAdd("M", -3, dtfecfin.Value) & " AL " & Me.dtfecfin.Value & " '"
    Rpt.Formulas(1) = "PER1 = '" & DateAdd("M", -2, dtfecfin.Value) & "'"
    Rpt.Formulas(2) = "PER2 = '" & DateAdd("M", -1, dtfecfin.Value) & "'"
    Rpt.Formulas(3) = "PER3 = '" & dtfecfin.Value & "'"
    Rpt.WindowTitle = "VENTAS DEL " & DateAdd("M", -3, dtfecfin.Value) & " AL " & Me.dtfecfin.Value
    Rpt.Action = 1
    Exit Sub
End If
cMensaje = stb1.SimpleText
stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
stb1.Refresh
Rpt.Connect = cCadConex
Rpt.ReportFileName = App.Path & crpt
Rpt.WindowTitle = "Reporte de Existencias"
Rpt.Formulas(0) = "FORMSELEC = " & ccondrpt
Rpt.Formulas(1) = "ENCABEZADO = '" & cEnca & " '"
'MsgBox Rpt.Formulas(0)
Rpt.Action = 1
stb1.SimpleText = cMensaje
stb1.Refresh
Exit Sub
Error:
  MsgBox Err.Description
  Unload Me
End Sub

Private Sub ordenaexistencias()
' Macro1 Macro
' Macro grabada el 30/04/2001 por Moises Leon
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("B3").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("H:I").Select
    Selection.NumberFormat = "$#,##0.00"   'Formateo las columnas de precios
    Columns("J:J").Select
    Selection.NumberFormat = "0"           'formateo las columnas código de barras
    Range("D17").Select
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Range("G5").Select
    ActiveWorkbook.Save
    Range("A8").Select
    
    Range("A1:J1").Select  'Selecciona el rango y pongo autofiltro
    Selection.AutoFilter
    With Selection.Interior  'le pongo color de relleno a las celdas color gris
        .ColorIndex = 15
        .Pattern = xlSolid
    End With

    
End Sub

Public Sub SELECTFILE()
On Error GoTo Error:
Cmdlg.FileName = ""
 Cmdlg.CancelError = True
 Cmdlg.DialogTitle = "Nombre del Archivo de Excel"
 Cmdlg.Filter = "Archivos Excel (*.xls) | *.xls"
 Cmdlg.ShowOpen
 cArch = Cmdlg.FileName
 If cArch = "" Or IsNull(cArch) Then
    MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
    Exit Sub
 End If
 Exit Sub
Error:
  MsgBox "Es necesario Especificar un nombre de archivo", vbCritical
End Sub


Private Sub INVCOSTO()
Rpt.ReportFileName = App.Path & "\COSTOINV.Rpt"
Rpt.WindowTitle = "Reporte de Existencias Con costos"
Rpt.Action = 1
End Sub

Private Sub cmdPedido_Click()
'Pendientes de confirmar
On Error GoTo Error:
crpt = "\Pedidos.rpt"
If OptPedido(0).Value = True Then
   ccondrpt = "{PEDIDOS.p_situacion} = 0"
   cEnca = "PENDIENTES DE CONFIRMAR"
ElseIf OptPedido(1).Value = True Then
    ccondrpt = "{PEDIDOS.P_situacion} = 1  AND {PEDIDOS.P_recibido} = 0"
    cEnca = "PENDIENTES DE RECIBIR"
ElseIf OptPedido(2).Value = True Then
    ccondrpt = "{PEDIDOS.P_recibido} = 1"
    cEnca = "RECIBIDOS"
ElseIf OptPedido(3).Value = True Then
    'En el reporte de eficiencia se toman en cuenta aquellos traslados que son por pedido, que se hayan enviado, que no esten cancelados y que sean salidas.
    ccondrpt = "MONTH( {TRASLADOS.t_fecha} ) = " & Month(txtfecha.Text) & " AND DAY( {TRASLADOS.t_fecha} ) = " & Day(txtfecha.Text) & " AND {TRASLADOS.t_tipo} = 0  AND {TRASLADOS.t_enviado} = 1  AND ISNULL({TRASLADOS.t_motivocancela}) AND {TRASLADOS.t_entrada} = 0 AND {TFPRODUC.PAQUETES} <> 1"
    cEnca = "SURTIDOS EL DIA " & UCase(Format(txtfecha.Text, "LONG DATE"))
    crpt = "\EfiAbaCa.rpt"
    cresp = MsgBox("DESEAS VER EL REPORTE DETALLADO", vbYesNo + vbDefaultButton2 + vbQuestion)
    If cresp = vbYes Then
       ccondrpt = ccondrpt + " AND {@DIFEREN} > 0"
       Rpt.SectionFormat(0) = "DETAIL;T;F;F;X;X;X;X"
       Rpt.SectionFormat(1) = "GH1;F;X;X;X;X;X;X"
       Rpt.SectionFormat(2) = "GF1;F;X;X;X;X;X;X"
    Else
       Rpt.SectionFormat(0) = "DETAIL;F;F;F;X;X;X;X"
       Rpt.SectionFormat(1) = "GH1;T;X;X;X;X;X;X"
       Rpt.SectionFormat(2) = "GF1;T;X;X;X;X;X;X"
    End If
ElseIf OptPedido(4).Value = True Then  'Total de entradas por tienda (pedidos sugeridos y traslados)
     cmern = stb1.SimpleText
     stb1.SimpleText = Space(35) & "Espere un momento generando reporte"
     stb1.Refresh
     Dim rsrep As ADODB.Recordset
     Set rsrep = New ADODB.Recordset
     cn.Execute "delete llegadasprod"
     fecha1 = dtfecini.Value
     fecha2 = dtfecfin.Value
     'MsgBox "FEcha Inicial " & fecha1
     'MsgBox "Fecha Final " & f echa2
     CAD = "insert into llegadasprod(producto,costo,importe,sol,llega,pedido) " & _
           " select dt_producto,dt_costo,dt_importe,dt_cantidad,dt_cantidad,dt_clave from traslados, detalletraslado" & _
           " Where dt_clave = t_clave And t_entrada = 1 and T_ENVIADO = 1 And year(t_fecha) = " & Year(fecha1) & " and month(t_fecha) = " & Month(fecha1) & " and day(t_fecha) >= " & Day(fecha1) & " and year(t_fecha) = " & Year(fecha2) & " and month(t_fecha) = " & Month(fecha2) & " and day(t_fecha) <= " & Day(fecha2) & " and dt_cantidad > 0"
     cn.Execute CAD
     CAD = "insert into llegadasprod(producto,costo,importe,sol,llega,pedido) " & _
           " select df_prod, df_costo, df_costo * df_cantreal, df_cantsol,df_cantreal, df_pedido from pedidos , detallefactura " & _
           " Where p_pedido = df_pedido And month(p_fecentreal) = " & Month(fecha1) & " and year(p_fecentreal) = " & Year(fecha1) & " and day(p_fecentreal) >= " & Day(fecha1) & " and year(p_fecentreal) = " & Year(fecha2) & " and  month(p_fecentreal) = " & Month(fecha2) & " and day(p_fecentreal) <=  " & Day(fecha2) & " and df_cantreal > 0 AND P_RECIBIDO = 1 "
     'MsgBox cad
     cn.Execute CAD
     Rpt.WindowTitle = "Total de entradas a Tienda  Bodega y Piso  "
     Rpt.ReportFileName = App.Path & "\entprod.rpt"
     encab = "TOTAL DE ENTRADAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
     Rpt.Formulas(0) = "ENCAB = '" & encab & "'"
     Rpt.Action = 1
     stb1.SimpleText = cmen
     stb1.Refresh
     Exit Sub
     Rpt.WindowTitle = "Total de entradas por Recibo"
     Rpt.ReportFileName = App.Path & "\todent.rpt"
     encab = "TOTAL DE ENTRADAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
     Rpt.Formulas(0) = "ENCAB = '" & encab & "'"
     cn.Execute "DELETE FROM llegadastmp"
     cn.Execute "INSERT INTO llegadastmp(pedido,fecha,cantsol,cantrec,impfac1,impfac2,pedprove) SELECT p_pedido,max(p_fecped),sum(df_cantsol),sum(df_cantreal),sum(df_cantreal*df_costo), venta = 0, tipo = 0 FROM pedidos,detallefactura WHERE df_pedido = p_pedido AND p_recibido = 1 AND p_cancelado = 0 AND p_fecentreal >= '" & dtfecini.Value & "' AND p_fecentreal <= '" & dtfecfin.Value & "' GROUP BY p_pedido"
     cn.Execute "INSERT INTO llegadastmp(pedido,fecha,cantsol,cantrec,impfac1,impfac2,pedprove) SELECT t_clave,max(t_fecha),sum(dt_cantidad),sum(dt_cantidad),sum(dt_cantidad*dt_costo), sum(dt_cantidad*dt_venta),tipo = 1 FROM traslados,detalletraslado WHERE dt_clave = t_clave AND t_enviado = 1 AND t_motivocancela IS NULL AND t_entrada = 1 AND t_fecha >= '" & Me.dtfecini.Value & "' AND t_fecha <= '" & dtfecfin.Value & "' GROUP BY t_clave"
     Rpt.Action = 1
     stb1.SimpleText = cmen
     stb1.Refresh
     Exit Sub
End If
If Trim(cmbProv.Text) <> "" Then ccondrpt = ccondrpt + " AND {PEDIDOS.p_proveedor} = '" & Trim(Mid(cmbProv.Text, Len(cmbProv.Text) - 5)) & "'"

cMensaje = stb1.SimpleText
stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
stb1.Refresh
Rpt.Connect = cCadConex
Rpt.ReportFileName = App.Path & crpt
Rpt.WindowTitle = "Reporte de pedidos por tienda"
Rpt.Formulas(0) = "FORMSELEC = " & ccondrpt
Rpt.Formulas(1) = "PEDIDO = 'LISTADO DE PEDIDOS POR TIENDA " & cEnca & " '"
'MsgBox Rpt.Formulas(0)
Rpt.Action = 1
stb1.SimpleText = cMensaje
stb1.Refresh
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Pedprov()
Dim rs As ADODB.Recordset
Dim crpt As String
On Error GoTo Error
crpt = "\PedProv.rpt"
If optPedpro(0).Value = True Then
   ccondrpt = "{PEDIDOS.p_situacion} = 0"
   cEnca = "PEDIDOS PENDIENTES DE CONFIRMAR"
   crpt = "\Pedidos.rpt"
ElseIf optPedpro(1).Value = True Then
    cEnca = "PEDIDOS PENDIENTES DE RECIBIR POR PROVEEDOR"
    ccondrpt = "{PEDPROVE.Pp_confirma} = 1  AND {PEDPROVE.Pp_recibe} = 0 AND {PEDPROVE.Pp_cancelado} = 0"
ElseIf optPedpro(2).Value = True Then
    cEnca = "PEDIDOS RECIBIDOS POR PROVEEDOR "
    ccondrpt = "{PEDPROVE.Pp_recibe} = 1"
ElseIf optPedpro(3).Value = True Then
    crpt = "\EfiProv.rpt"
    cEnca = "LLEGADAS DEL " & Me.dtfecini & " AL " & Me.dtfecfin & IIf(Trim(cmbProvPr.Text) <> "", " DE " & Trim(cmbProvPr.Text), "")
    If Trim(cmbprod.Text) <> "" Then
       Rpt.SectionFormat(0) = "DETAIL;T;F;F;X;X;X;X"
       Rpt.SectionFormat(1) = "GH2;T;F;F;X;X;X;X"
       Rpt.SectionFormat(2) = "GF2;T;F;F;X;X;X;X"
       ccondrpt = "{PEDPROVE.pp_recibe} = 1 AND {PEDPROVE.pp_fecrecibe}  >= Date( " & Format(dtfecini.Value, "yyyy,mm,dd") & ") AND {PEDPROVE.pp_fecrecibe} <= DATE(" & Format(dtfecfin.Value, "yyyy,mm,dd") & ") AND {TFPRODUC.CONSEC} = '" & Trim(Mid(cmbprod.Text, Len(cmbprod.Text) - 10)) & "'"
    Else  'No se ha seleccionado producto
       If Trim(cmbProvPr.Text) <> "" Then
          Set rs = New ADODB.Recordset
          rs.Open "SELECT DG_PRODUCTO,SUM(DG_CANTSOL), SUM(DG_CANTREAL) FROM DETALLEGLOBAL,PEDPROVE WHERE DG_PEDIDO = PP_PEDIDO AND PP_PROVEEDOR = '" & Trim(Mid(cmbProvPr.Text, Len(cmbProvPr.Text) - 5)) & "' AND PP_FECRECIBE >= '" & Me.dtfecfin & "' AND PP_FECRECIBE <= DATEADD(day, 1, '" & dtfecfin.Value & "' ) GROUP BY DG_PRODUCTO HAVING SUM(DG_CANTREAL) = 0", cn, adOpenKeyset, adLockOptimistic, adCmdText
          Rpt.Formulas(2) = "VARNOSUR = '" & CStr(rs.RecordCount) & "'"
          Rpt.SectionFormat(0) = "DETAIL;T;F;F;X;X;X;X"
          Rpt.SectionFormat(1) = "GH2;T;F;F;X;X;X;X"
          Rpt.SectionFormat(2) = "GF2;T;F;F;X;X;X;X"
       Else
          Rpt.SectionFormat(0) = "DETAIL;F;F;F;X;X;X;X"
          Rpt.SectionFormat(1) = "GH2;F;F;F;X;X;X;X"
          Rpt.SectionFormat(2) = "GF2;F;F;F;X;X;X;X"
       End If
       ccondrpt = "{PEDPROVE.pp_recibe} = 1 AND {PEDPROVE.pp_fecrecibe}  >= Date( " & Format(dtfecini.Value, "yyyy,mm,dd") & ") AND {PEDPROVE.pp_fecrecibe} <= DATE(" & Format(dtfecfin.Value, "yyyy,mm,dd") & ") AND ( {DETALLEGLOBAL.dg_cantsol} > 0 OR {DETALLEGLOBAL.dg_cantsolp} > 0 )"
    End If
ElseIf optPedpro(4).Value = True Then
    crpt = "\EfiProvC.rpt"
    If Trim(cmbprod.Text) <> "" Then
       ccondrpt = "{PEDPROVE.pp_recibe} = 1 AND {PEDPROVE.pp_fecrecibe}  >= Date( " & Format(dtfecini.Value, "yyyy,mm,dd") & ") AND {PEDPROVE.pp_fecrecibe} <= DATE(" & Format(dtfecfin.Value, "yyyy,mm,dd") & ") AND UPPERCASE({CATPROV.comprador}) = '" & Trim(cmbprod.Text) & "' AND ( {DETALLEGLOBAL.dg_cantsol} > 0 OR {DETALLEGLOBAL.dg_cantsolp} > 0 )"
    Else
       ccondrpt = "{PEDPROVE.pp_recibe} = 1 AND {PEDPROVE.pp_fecrecibe}  >= Date( " & Format(dtfecini.Value, "yyyy,mm,dd") & ") AND {PEDPROVE.pp_fecrecibe} <= DATE(" & Format(dtfecfin.Value, "yyyy,mm,dd") & ") AND ( {DETALLEGLOBAL.dg_cantsol} > 0 OR {DETALLEGLOBAL.dg_cantsolp} > 0 )"
    End If
    cEnca = "LLEGADAS DE PROVEEDORES DEL " & dtfecini.Value & " AL " & dtfecfin.Value & IIf(Trim(cmbprod.Text) <> "", " DEL COMPRADOR (A) " & Trim(cmbprod.Text), "")
ElseIf optPedpro(5).Value = True Then
    Rpt.ReportFileName = App.Path & "\pecompra.rpt"
    cadsql = "SELECT PEDPROVE.pp_proveedor, PEDPROVE.pp_fechagen, PEDPROVE.pp_recibe, PEDPROVE.pp_fecrecibe, " & _
                     "DETALLEGLOBAL.dg_cantreal, DETALLEGLOBAL.dg_promocionr, " & _
                     "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & _
             "FROM PITICO.dbo.PEDPROVE PEDPROVE, " & _
                     "PITICO.dbo.DETALLEGLOBAL DETALLEGLOBAL, " & _
                     "PITICO.dbo.TFPRODUC TFPRODUC " & _
             "WHERE PEDPROVE.pp_pedido = DETALLEGLOBAL.dg_pedido AND " & _
                     "DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND " & _
                     "PEDPROVE.pp_proveedor = 'C52' AND PEDPROVE.pp_recibe = 1 AND " & _
                      "PEDPROVE.pp_fecrecibe >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND PEDPROVE.pp_fecrecibe  <= '" & Format(dtfecfin.Value, "yyyy-dd-mm") & "'"
    Rpt.Action = 1
Else
    Call PEDPROVE   'Rutina para reportes del areea de compras
    Exit Sub
End If
If Trim(cmbProvPr.Text) <> "" And optPedpro(0).Value = False Then
   ccondrpt = ccondrpt + " AND {PEDPROVE.pp_proveedor} = '" & Trim(Mid(cmbProvPr.Text, Len(cmbProvPr.Text) - 5)) & "'"
ElseIf optPedpro(0).Value = True And Trim(cmbProvPr.Text) <> "" Then
   ccondrpt = ccondrpt + " AND {PEDIDOS.p_proveedor} = '" & Trim(Mid(cmbProvPr.Text, Len(cmbProvPr.Text) - 5)) & "'"
End If

cMensaje = stb1.SimpleText
stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
stb1.Refresh
Rpt.Connect = cCadConex
Rpt.ReportFileName = App.Path & crpt
Rpt.WindowTitle = "Reporte de pedidos por proveedor"
Rpt.Formulas(0) = "FORMSELEC = " & ccondrpt
Rpt.Formulas(1) = "PEDIDO = '" & cEnca & "'"
'MsgBox Rpt.Formulas(0)
Rpt.Action = 1
stb1.SimpleText = cMensaje
stb1.Refresh
Exit Sub
Error:
  MsgBox Err.Description
  Unload Me
End Sub

Private Sub reptraslados()
Dim ccondrpt
Dim fecini As Date
Dim fecfin As Date
On Error GoTo Error:
fecini = dtfecini.Value
fecfin = dtfecfin.Value
If OptTrasl(0).Value Then
    'cn.Execute " reptraslados  '" & dtfecini.Value & "' , '" & dtfecfin.Value & "'"
    crpt = "\contrasl.rpt"
    cencab = "TRASLADOS ENVIADOS A TIENDAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    Rpt.WindowTitle = cencab
    Rpt.Connect = cCadConex
    Rpt.ReportFileName = App.Path & crpt
    Rpt.Formulas(0) = "ENCAB = '" & cencab & "'"
    Rpt.Formulas(1) = "FORMSELEC = {TRASLADOS.t_fecha} >= DATE(" & Format(Me.dtfecini.Value, "YYYY,MM,DD") & ") AND {TRASLADOS.t_fecha} <= DATE(" & Format(Me.dtfecfin.Value, "YYYY,MM,DD") & _
                      ") AND ISNULL({TRASLADOS.t_motivocancela}) AND {TRASLADOS.t_entrada} = 0 AND {TRASLADOS.t_enviado} = 1"
    Rpt.Action = 1
    stb1.SimpleText = cMensaje
    stb1.Refresh
    Exit Sub
ElseIf OptTrasl(1).Value Then
    crpt = "\InvSinde.rpt"
    cencab = "PRODUCTOS CON EXISTENCIA Y SIN DESPLAZAMIENTO DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    cCadena = "SELECT INVENTARIO.inprod, INVENTARIO.incant, INVENTARIO.incantpza," & _
                      "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES " & _
              "FROM PITICO.dbo.INVENTARIO INVENTARIO, " & _
                    "PITICO.dbo.TFPRODUC TFPRODUC " & _
              "WHERE INVENTARIO.inprod NOT IN " & _
                    "(SELECT DETALLETRASLADO.DT_producto " & _
                    "FROM PITICO.dbo.DETALLETRASLADO DETALLETRASLADO,  PITICO.dbo.TRASLADOS TRASLADOS " & _
                    "WHERE DETALLETRASLADO.dt_clave =  TRASLADOS.t_clave AND TRASLADOS.t_enviado = 1 AND TRASLADOS.t_entrada = 0 AND TRASLADOS.t_motivocancela IS NULL AND TRASLADOS.t_fecha >= '" & Format(fecini, "yyyy-dd-mm") & "' AND TRASLADOS.t_fecha <= '" & Format(DateAdd("d", 1, fecfin), "yyyy-dd-mm") & "') " & _
                           "AND (INVENTARIO.incant > 0 OR INVENTARIO.incantpza > 0 )AND INVENTARIO.inprod  = TFPRODUC.consec"
ElseIf OptTrasl(2).Value Then
    crpt = "\contrasl.rpt"
    fecini = fecini - 1
    fecfin = fecfin + 1
    ccondrpt = "{TRASLADOS.t_fecha} > DATE(" & Format(fecini, "yyyy,mm,dd") & ") AND {TRASLADOS.t_fecha} < DATE( " & Format(fecfin, "yyyy,mm,dd") & ") AND ISNULL({TRASLADOS.t_motivocancela})  AND {TRASLADOS.t_entrada} = 0 and {TRASLADOS.t_enviado} = 1 "
    cencab = "TRASLADOS PENDIENTES DE ENVIAR A TIENDAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    cCadena = "SELECT TRASLADOS.t_clave, TRASLADOS.t_fecha, TRASLADOS.t_tipo, TRASLADOS.t_enviado, TRASLADOS.t_motivocancela, TRASLADOS.t_entrada, TRASLADOS.t_foliotie,TRASLADOS.t_papeleria," & _
                       "DETALLETRASLADO.dt_producto, DETALLETRASLADO.dt_cantidad, DETALLETRASLADO.dt_cantidadp, DETALLETRASLADO.dt_costo, " & _
                       "CATTIENDA.ticlave, CATTIENDA.tidescrip, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES " & Chr(13) & _
               "FROM   pitico.dbo.TRASLADOS TRASLADOS, " & _
                       "pitico.dbo.DETALLETRASLADO DETALLETRASLADO, " & _
                       "pitico.dbo.CATTIENDA CATTIENDA," & _
                       "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
               "WHERE  TRASLADOS.t_clave = DETALLETRASLADO.dt_clave AND " & _
                       "TRASLADOS.t_sucursalreceptor = CATTIENDA.ticlave AND " & _
                       "TRASLADOS.t_fecha >'" & Format(fecini, "yyyy-dd-mm") & "' AND TRASLADOS.t_fecha < '" & Format(fecfin, "yyyy-dd-mm") & "'  AND " & _
                       "DETALLETRASLADO.dt_producto = TFPRODUC.CONSEC AND " & _
                       "TRASLADOS.t_enviado = 0 AND TRASLADOS.t_entrada = 0 AND TRASLADOS.t_sucursalreceptor <> '81' "
    If chkvolumen.Value = 1 Then
        cCadena = cCadena & " AND TRASLADOS.t_observa = 'VOLUMEN' "
        cencab = "TRASLADOS DE VOLUMEN PENDIENTES DE ENVIAR  A TIENDAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    End If
    cCadena = cCadena & Chr(13) & "ORDER BY  CATTIENDA.tidescrip ASC, TRASLADOS.t_fecha ASC, TRASLADOS.t_clave ASC"
ElseIf OptTrasl(3).Value Then
    'If mskhoraIni.Text >= mskHoraFin.Text Then
    '   MsgBox "LA HORA FINAL DEBE SER MAYOR A LA INCIAL", vbInformation
    '   mskHoraFin.SetFocus
    '   Exit Sub
    'End If
    cn.Execute "DELETE FROM tragon"
    If MsgBox("Deseas ver el reporte a total por día", vbQuestion + vbYesNo) = vbYes Then
       cn.Execute "INSERT INTO tragon(usuario,clave,fecha,cajas,piezas,varprod,costo,ventas) SELECT name,max(t_clave),CONVERT(CHAR,T_FECHA,5), sum(dt_cantidad), sum(dt_cantidad * paquetes) + sum(dt_cantidadp), count(*), sum(dt_cantidad* dt_costo) + sum(dt_cantidadp*dt_costop), sum(dt_cantidad* dt_venta) + sum(dt_cantidadp*dt_ventap) FROM usuarios,traslados,detalletraslado,tfproduc WHERE consec = dt_producto and clave =* t_gon and t_gon <> 0 and dt_clave = t_clave and T_MOTIVOCANCELA IS NULL AND t_fecha >= '" & dtfecini.Value & " " & mskhoraIni.Text & "' AND t_fecha <= '" & dtfecfin.Value & " " & mskHoraFin.Text & "' GROUP BY name,CONVERT(CHAR,T_FECHA,5)"
    Else
       cn.Execute "INSERT INTO tragon(usuario,clave,fecha,cajas,piezas,varprod,costo,ventas) SELECT name,t_clave,max(T_FECHA), sum(dt_cantidad), sum(dt_cantidad * paquetes) + sum(dt_cantidadp), count(*), sum(dt_cantidad* dt_costo) + sum(dt_cantidadp*dt_costop), sum(dt_cantidad* dt_venta) + sum(dt_cantidadp*dt_ventap) FROM usuarios,traslados,detalletraslado,tfproduc WHERE consec = dt_producto and clave = t_gon and t_gon <> 0 and dt_clave = t_clave and T_MOTIVOCANCELA IS NULL AND t_fecha >= '" & dtfecini.Value & " " & mskhoraIni.Text & "' AND t_fecha <= '" & dtfecfin.Value & " " & mskHoraFin.Text & "' GROUP BY name,T_clave"
    End If
    Rpt.WindowTitle = "Traslados por gondolero"
    Rpt.ReportFileName = App.Path & "\tragon.rpt"
    Rpt.Formulas(0) = "ENCABEZADO = 'ENVIOS A PISO DEL " & dtfecini.Value & " AL " & dtfecfin.Value & " EN EL HORARIO: " & mskhoraIni.Text & " - " & mskHoraFin.Text & " HRS.'"
    Rpt.Action = 1
    Exit Sub
ElseIf OptTrasl(4).Value Then
    Rpt.WindowTitle = "Salida de productos en Traslados"
    Rpt.ReportFileName = App.Path & "\contrasd.rpt"
    Rpt.Formulas(0) = "ENCAB = 'SALIDA DE PRODUCTOS A TRAVES DE TRASLADOS DEL " & Me.dtfecini.Value & " AL " & Me.dtfecfin.Value & "'"
    Rpt.SQLQuery = "SELECT TRASLADOS.t_fecha, " & _
                           "DETALLETRASLADO.dt_producto, DETALLETRASLADO.dt_cantidad, DETALLETRASLADO.dt_cantidadp, " & _
                           "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES " & Chr(13) & _
                   "FROM pitico.dbo.TRASLADOS TRASLADOS, " & _
                           "pitico.dbo.DETALLETRASLADO DETALLETRASLADO," & _
                           "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                   "WHERE TRASLADOS.t_clave = DETALLETRASLADO.dt_clave AND " & _
                           "DETALLETRASLADO.dt_producto = TFPRODUC.CONSEC AND " & _
                           "TRASLADOS.t_fecha >= '" & Format(Me.dtfecini.Value, "yyyy-dd-mm") & "' AND TRASLADOS.t_fecha <= '" & Format(DateAdd("d", 1, dtfecfin.Value), "yyyy-dd-mm") & "' AND " & _
                           "TRASLADOS.t_enviado = 1 AND TRASLADOS.t_entrada = 0 " & Chr(13) & _
                   "ORDER BY DETALLETRASLADO.dt_producto ASC "
    Rpt.Action = 1
    Exit Sub
End If
Rpt.WindowTitle = LCase(cencab)
Rpt.SQLQuery = cCadena
cMensaje = stb1.SimpleText
stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
stb1.Refresh
Rpt.Connect = cCadConex
Rpt.ReportFileName = App.Path & crpt
Rpt.Formulas(0) = "FORMSELEC = " & ccondrpt
Rpt.Formulas(1) = "ENCAB = '" & cencab & "'"
Rpt.Action = 1
stb1.SimpleText = cMensaje
stb1.Refresh
Exit Sub
Error:
  MsgBox Err.Description
  Unload Me
End Sub

Private Sub rptVentas()
Dim cCond As String
Dim lmismatda As Boolean
Dim rs As ADODB.Recordset
On Error GoTo Error:
cMensaje = stb1.SimpleText
stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
stb1.Refresh
Rpt.SectionFormat(0) = ""
Rpt.SectionFormat(1) = ""
Rpt.SectionFormat(2) = ""
Rpt.SectionFormat(3) = ""
Rpt.SectionFormat(4) = ""
Rpt.SectionFormat(5) = ""
Rpt.SectionFormat(6) = ""
Rpt.Formulas(0) = ""
Rpt.Formulas(1) = ""
Rpt.SortFields(0) = ""
Rpt.DataFiles(0) = ""
If OptVentas(1).Value Then
   cmensa = stb1.SimpleText
   Me.stb1.SimpleText = Space(25) & "Calculando ventas a contado del " & Me.dtfecini.Value & " al " & Me.dtfecfin.Value
   stb1.Refresh
   If Trim(Mid(cSucursal, 1, 2)) = "16" Then
      contado = "Y1"
      credito = " F.SERIE IN ('B','CCC') "
   ElseIf Trim(Mid(cSucursal, 1, 2)) = "23" Then
      contado = "D2"
   ElseIf Trim(Mid(cSucursal, 1, 2)) = "24" Then
      contado = "I2"
      credito = "J2"
   ElseIf Trim(Mid(cSucursal, 1, 2)) = "28" Then
      contado = "JJJ"
      credito = " F.SERIE IN ('LLL','KKK') "
   ElseIf Trim(Mid(cSucursal, 1, 2)) = "28" Then
      contado = "G2"
      credito = " F.SERIE IN ('H2','DDD') "
   ElseIf Trim(Mid(cSucursal, 1, 2)) = "13" Then
      contado = "AB"
      credito = " F.SERIE IN ('ABX') "
   End If
   cn.Execute "DELETE FROM VtaConcre"
   cn.Execute "INSERT INTO VTACONCRE(vtacon,vtacre,fecha,vtacrepen) SELECT SUM(D.IMPORTE) AS VTACON, CREDITO = 0, D.FECHA_det, CREPEN = 0 FROM FACVENTA_DET D WHERE D.fecha_det >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND d.fecha_det <= '" & Format(dtfecfin.Value, "yyyy-dd-mm") & "' AND SERIE = '" & contado & "' GROUP BY d.FECHA_det ORDER BY d.fecha_det"
   Set rs = New ADODB.Recordset
   rs.Open "SELECT * FROM VTACONCRE", cn, adOpenDynamic, adLockOptimistic, adCmdText
   Set RSTEMP = New ADODB.Recordset
   vtacon = 0: vtacre = 0: vtacrepen = 0: vtacref = 0
   While Not rs.EOF
      stb1.SimpleText = Space(25) & "Calculando ventas a credito del dia " & Format(rs!fecha, "long date")
      stb1.Refresh
      RSTEMP.Open "SELECT SUM(F.TOTAL) AS VTACRE FROM FACVENTA F WHERE F.FACFECHA >= '" & Format(rs!fecha, "yyyy-dd-mm") & "' AND F.FACFECHA <= '" & Format(rs!fecha, "yyyy-dd-mm") & "' AND " & credito & " AND F.cancelado = 0", cn, adOpenDynamic, adLockOptimistic, adCmdText
      rs!vtacref = IIf(IsNull(RSTEMP!vtacre), 0, RSTEMP!vtacre)
      rs.Update: RSTEMP.Close
      RSTEMP.Open "SELECT SUM(F.porpagar) AS VTACREPEN FROM FACVENTA F WHERE f.FacFECHA >= '" & Format(rs!fecha, "yyyy-dd-mm") & "' AND f.facFECHA <= '" & Format(rs!fecha, "yyyy-dd-mm") & "' AND " & credito & " AND F.COBRADO = 0 AND f.cancelado = 0", cn, adOpenDynamic, adLockOptimistic, adCmdText
      rs!vtacrepen = IIf(IsNull(RSTEMP!vtacrepen), 0, RSTEMP!vtacrepen)
      rs!vtacre = rs!vtacref - IIf(IsNull(RSTEMP!vtacrepen), 0, RSTEMP!vtacrepen)
      vtacon = vtacon + rs!vtacon
      vtacref = vtacref + rs!vtacref
      vtacre = vtacre + rs!vtacre
      vtacrepen = vtacrepen + rs!vtacrepen
      rs!avtacon = vtacon
      rs!avtacref = vtacref
      rs!avtacre = vtacre
      rs!avtacrepen = vtacrepen
      rs.Update: RSTEMP.Close
      rs.MoveNext
   Wend
   Set rs = Nothing
   crpt = "\VtaconCRE.rpt"
   cencab = "RELACION DE FACTURAS A CONTADO Y CREDITO DEL " & Me.dtfecini.Value & " AL " & Me.dtfecfin.Value
   Rpt.SQLQuery = ""
   Rpt.SectionFormat(0) = "":    Rpt.SectionFormat(1) = ""
   Rpt.SectionFormat(2) = "":    Rpt.SectionFormat(3) = ""
   Rpt.SectionFormat(4) = "":    Rpt.SectionFormat(5) = ""
   Rpt.SectionFormat(6) = ""
   stb1.SimpleText = cmensa
   stb1.Refresh
ElseIf OptVentas(2).Value Then
   stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
   stb1.Refresh
   cn.Execute "DELETE FROM VtaConcre"
   cn.Execute "INSERT INTO VTACONCRE(fecha,vtacon,vtacre,J2) SELECT FECHA_det, SUM( CASE serie WHEN 'Y1' THEN importe END) AS Y1, SUM( CASE  WHEN serie IN ('B','CCC') THEN importe END) AS B, SUM( CASE serie WHEN 'J2' THEN importe END) AS DIF " & _
              "FROM facventa_det, tfproduc WHERE consec = producto AND fecha_det >= '" & dtfecini.Value & "' AND fecha_det <= '" & dtfecfin.Value & "' GROUP BY fecha_det "
   cn.Execute "UPDATE vtaconcre SET vtacre = 0 WHERE vtacre IS NULL"
   Set rs = New ADODB.Recordset
   rs.Open "SELECT SUM(importe) AS Vta, fecha_det, serie FROM facvtamay WHERE fecha_det >= '" & Me.dtfecini.Value & _
           "' AND fecha_det <= '" & dtfecfin.Value & "' GROUP BY fecha_det, serie ORDER BY FECHA_DET", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
   Dim RST As ADODB.Recordset
   Set RST = New ADODB.Recordset
   While Not rs.EOF
      RST.Open "SELECT * FROM vtaconcre WHERE fecha = '" & rs!fecha_det & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
      If RST.BOF And RST.EOF Then
         cn.Execute "INSERT INTO vtaconcre(fecha,vtacon,vtacre) VALUES ('" & rs!fecha_det & "',0,0)"
      End If
      'A partir del 20 de Abril se Iniciaron Actividades en Cosijopi.
      Dim FechaIni As Date: Dim FechaFin As Date
      'FechaIni = "19/04/02": FechaFin = "20/05/02"   COSIJOPI
      FechaIni = "26/05/03": FechaFin = "04/06/03"   'MIAHUATLAN
      'If (rs!fecha_det >= FechaIni And rs!fecha_det <= FechaFin) And rs!SERIE = "D2" Then
      '   'Modulo N
      '   cn.Execute "UPDATE vtaconcre SET D2 = (SELECT SUM(IMPORTE) FACTURA FROM facvtamay WHERE SERIE = 'D2' AND fecha_det = '" & rs!fecha_det & "' AND (CAST(factura AS MONEY) <  11100 OR (CAST(factura AS MONEY) BETWEEN  11301 AND 11500) OR (CAST(factura AS MONEY) BETWEEN  11601 AND 11699) OR (CAST(factura AS MONEY) BETWEEN  11901 AND 12082) OR CAST(factura AS MONEY) > 12091) ) WHERE fecha = '" & rs!fecha_det & "'"
      '   'Cosijopi
      '   cn.Execute "UPDATE vtaconcre SET I2 = (SELECT SUM(IMPORTE) FACTURA FROM facvtamay WHERE SERIE = 'D2' AND fecha_det = '" & rs!fecha_det & "' AND ((CAST(factura AS MONEY) BETWEEN  11100 AND 11300) OR (CAST(factura AS MONEY) BETWEEN  11501 AND 11600) OR (CAST(factura AS MONEY) BETWEEN  11700 AND 11900) OR (CAST(factura AS MONEY) BETWEEN 12083 AND 12090) )) WHERE fecha = '" & rs!fecha_det & "'"
      If (rs!fecha_det >= FechaIni And rs!fecha_det <= FechaFin) And Trim(rs!SERIE) = "B" Then
         cn.Execute "UPDATE vtaconcre SET GGG = " & IIf(IsNull(rs!vta), 0, rs!vta) & " WHERE fecha = '" & rs!fecha_det & "'"
      Else
         If Trim(rs!SERIE) = "I2" Then
            cn.Execute "UPDATE vtaconcre SET D2 = D2 + " & IIf(IsNull(rs!vta), 0, rs!vta) & " WHERE fecha = '" & rs!fecha_det & "'"
         Else
            cn.Execute "UPDATE vtaconcre SET " & rs!SERIE & "= " & rs!SERIE & " + " & IIf(IsNull(rs!vta), 0, rs!vta) & " WHERE fecha = '" & rs!fecha_det & "'"
         End If
      End If
      rs.MoveNext
      RST.Close
   Wend
   cn.Execute "UPDATE vtaconcre SET vtacon = 0 WHERE vtacon IS NULL"
   cn.Execute "UPDATE vtaconcre SET I2 = 0 WHERE I2 IS NULL"
   cn.Execute "UPDATE vtaconcre SET J2 = 0 WHERE J2 IS NULL"
   cn.Execute "UPDATE vtaconcre SET DDD = 0 WHERE DDD IS NULL"
   cn.Execute "UPDATE vtaconcre SET D = 0 WHERE D IS NULL"
   crpt = "\todmay.rpt"
   cencab = "VENTAS POR BODEGA DEL " & dtfecini.Value & " AL " & dtfecfin.Value
   Rpt.SQLQuery = ""
ElseIf OptVentas(0).Value Then
    Rpt.WindowTitle = "Reporte de Ventas por producto"
    cCond = ""
    nsepar = InStr(1, cmbprod.Text, "|")
    If nsepar > 0 Then
        cCond = " AND producto = '" & Trim(Mid(cmbprod.Text, nsepar + 1)) & "' "
        cencab = "DESPLAZAMIENTO DE " & Trim(cmbprod.Text)
    ElseIf cmbProVta.Text <> "" Then
        nsepar = InStr(1, cmbProVta.Text, "|")
        cCond = " AND CLAPROVE = '" & Trim(Mid(cmbProVta.Text, nsepar + 1)) & "' "
        cencab = "DESPLAZAMIENTO DE " & Trim(cmbProVta.Text)
    Else
        cencab = "DESPLAZAMIENTO DE VENTAS "
    End If
    If OptSerie(0).Value = True Then
        'SUCUR = "BODEGA MIGUEL CABRERA"
        'lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(0).Caption, 1, 2))
        'campo = "CAB"
        'cCond = cCond + " AND serie IN ('Y1','B','CCC')"
        SUCUR = "PITICO13"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(0).Caption, 1, 2))
        campo = "CAB"
        cCond = cCond + " AND serie IN ('AB')"

    ElseIf OptSerie(1).Value = True Then
        SUCUR = "BODEGA CENTRAL DE ABASTOS"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(1).Caption, 1, 2))
        campo = "CEN"
        cCond = cCond + " AND serie = 'D2'"
    ElseIf OptSerie(2).Value = True Then
        SUCUR = "SUCURSAL PUERTO ESCONDIDO"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(2).Caption, 1, 2))
        campo = "PTO"
        cCond = cCond + " AND serie IN ('G2','H2','DDD')"
    ElseIf OptSerie(3).Value = True Then
        SUCUR = "BODEGA MIAHUATLAN"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(3).Caption, 1, 2))
        campo = "MIA"
        cCond = cCond + " AND serie IN ('GGG','HHH')"
    ElseIf OptSerie(5).Value = True Then
        SUCUR = "SUCURSAL ISTMO"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(5).Caption, 1, 2))
        campo = "IST"
        cCond = cCond + " AND serie IN ('JJJ','KKK','LLL')"
    End If
    cn.Execute "DELETE FROM desplatod"
    cadsql = ""
        cTipOrd = IIf(OptOrden(0).Value = True, "+", "-")
        If OptTipVta(0).Value = True Then
            Rpt.SortFields(0) = cTipOrd & "{TFPRODUC.DESCRIPC}"
            Rpt.GroupSortFields(0) = cTipOrd & "COUNT ({TFPRODUC.CONSEC}, {TFPRODUC.CLAPROVE})"
        ElseIf OptTipVta(1).Value = True Then
            Rpt.SortFields(0) = cTipOrd & "{@TOTAL}"
            Rpt.GroupSortFields(0) = cTipOrd & "Sum ({@TOTAL}, {TFPRODUC.CLAPROVE})"
        ElseIf OptTipVta(2).Value = True Then
            Rpt.SortFields(0) = cTipOrd & "{@IMPORTE}"
            Rpt.GroupSortFields(0) = cTipOrd & "Sum ({@IMPORTE}, {TFPRODUC.CLAPROVE})"
        ElseIf OptTipVta(3).Value = True Then
            Rpt.SortFields(0) = cTipOrd & "{@UTILIDAD}"
            Rpt.GroupSortFields(0) = cTipOrd & "Sum ({@UTILIDAD}, {TFPRODUC.CLAPROVE})"
        ElseIf OptTipVta(4).Value = True Then
            Rpt.SortFields(0) = cTipOrd & "{@PMARGEN}"
            Rpt.GroupSortFields(0) = cTipOrd & "sum ({@PMARGEN}, {TFPRODUC.CLAPROVE})"
        End If

    If OptSerie(4).Value = True Then  'Desplazamiento de todas las bodegas de mayoreo
        'Desplazamientos de Matriz
        cn.Execute "INSERT INTO desplatod(Consec,cab,costo,importe,utilidad) SELECT producto, SUM(CANTIDAD), SUM( ((cantidad * costo) + (cantidadp * costop)) ) ,SUM(IMPORTE), SUM( importe - ((cantidad * costo) + (cantidadp * costop)) ) " & _
               "FROM facventa_det, tfproduc WHERE importe > 0 and producto = consec AND fecha_det >= '" & dtfecini.Value & "' AND fecha_det <= '" & Me.dtfecfin.Value & "' " & cCond & _
               "GROUP BY producto"
        Set rs = New ADODB.Recordset
        'Desplazamientos de sucursales
        rs.Open "SELECT producto, PTO = CASE WHEN serie = 'G2' OR serie = 'H2' OR serie = 'DDD' THEN SUM(cantidad) END , " & _
               "CEN = CASE WHEN serie = 'D2' THEN SUM(CANTIDAD) END , " & _
               "IST = CASE WHEN serie = 'D' OR SERIE = 'JJJ' OR SERIE = 'KKK' OR SERIE = 'LLL'  THEN SUM(CANTIDAD) END , " & _
               "MIA = CASE WHEN SERIE = 'B' OR SERIE = 'GGG' OR SERIE = 'HHH' THEN SUM(CANTIDAD) END , CAB = 0, SUM(importe) AS importe, SUM( ((cantidad * costo) + (cantidadp * costop)) ) Costo, SUM( importe - ((cantidad * costo) + (cantidadp * costop)) ) Utilidad " & _
               "FROM facvtamay, tfproduc WHERE producto = consec AND importe > 0 AND fecha_det >= '" & Me.dtfecini.Value & "' AND fecha_det <= '" & Me.dtfecfin.Value & "' " & cCond & _
               "GROUP BY producto,serie ORDER BY PRODUCTO", cn, adOpenStatic, adLockOptimistic, adCmdText
        Set RSTEMP = New ADODB.Recordset
        While Not rs.EOF
            RSTEMP.Open "SELECT * FROM desplatod WHERE consec = '" & Trim(rs!producto) & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
            If RSTEMP.BOF And RSTEMP.EOF Then
               cn.Execute "INSERT INTO desplatod(consec) VALUES ('" & Trim(rs!producto) & "')"
            End If
            cn.Execute "UPDATE desplatod SET pto = pto +" & IIf(IsNull(rs!pto), 0, rs!pto) & ", CEN = cen +" & IIf(IsNull(rs!cen), 0, rs!cen) & _
                       ", MIA = MIA + " & IIf(IsNull(rs!MIA), 0, rs!MIA) & _
                       ", IST = IST + " & IIf(IsNull(rs!IST), 0, rs!IST) & _
                       ", importe = importe + " & rs!importe & ",costo = costo + " & IIf(IsNull(rs!costo), 0, rs!costo) & ",utilidad  = utilidad + " & IIf(IsNull(rs!UTILIDAD), 0, rs!UTILIDAD) & " WHERE consec  = '" & Trim(rs!producto) & "'"
'                       ", Cabp = Cabp +" & IIf(IsNull(RS!Ptop), 0, RS!Ptop) & "+" & IIf(IsNull(RS!Cenp), 0, RS!Cenp) & ", COS = cos +" & IIf(IsNull(RS!cosi), 0, RS!cosi)
            RSTEMP.Close
            rs.MoveNext
        Wend
        cn.Execute "UPDATE desplatod SET pto = 0 WHERE pto IS NULL"
        cn.Execute "UPDATE desplatod SET cen = 0 WHERE cen IS NULL"
        cn.Execute "UPDATE desplatod SET Mia = 0 WHERE Mia IS NULL"
        cn.Execute "UPDATE desplatod SET cab = 0 WHERE cab IS NULL"
        cn.Execute "UPDATE desplatod SET IST = 0 WHERE IST IS NULL"
        SUCUR = "DESPLAZAMIENTO DE TODAS LAS BODEGAS DE MAYOREO"
    Else
        cn.Execute "INSERT INTO desplatod(Consec," & campo & ",costo,importe,utilidad) SELECT producto, SUM(CANTIDAD), SUM( ((cantidad * costo) + (cantidadp * costop)) ) ,SUM(IMPORTE), SUM( importe - ((cantidad * costo) + (cantidadp * costop)) ) " & _
               "FROM " & IIf(lmismatda, "facventa_det", "FacVtaMay") & ", tfproduc WHERE importe > 0 and producto = consec AND fecha_det >= '" & dtfecini.Value & "' AND fecha_det <= '" & Me.dtfecfin.Value & "' " & cCond & _
               "GROUP BY producto"
    
    End If
    Rpt.Formulas(1) = "SUCURSAL = '" & SUCUR & "'"
    cencab = cencab & " DEL " & Me.dtfecini.Value & " AL " & Me.dtfecfin.Value
    If MsgBox("Deseas ver tambien reporte por proveedor", vbYesNo + vbQuestion + vbDefaultButton2, "Agrupado?") = vbYes Then
       crpt = "\desptod1.rpt"
       Rpt.GroupCondition(0) = "GROUP1;{TFPRODUC.CLAPROVE};ANYCHANGE;A"
       Rpt.Formulas(0) = "ENCAB = '" & cencab & "'"
       Rpt.ReportFileName = App.Path & crpt
       Rpt.WindowTitle = cencab
       Rpt.Action = 1
    End If
    crpt = "\desptod.rpt"
    Rpt.GroupCondition(0) = ""
    Rpt.GroupSortFields(0) = ""
ElseIf OptVentas(3).Value Then
    Dim diatra As Variant
    'diatra = Array(0, 26, 24, 26, 25, 27, 25, 27, 25, 26, 27, 25, 30)
    diatra = Array(0, 26, 24, 26, 28, 27, 25, 27, 25, 26, 27, 25, 30)
    Set RST = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    cn.Execute "DELETE FROM vtaconcre"
    If OptSerie(0).Value = True Then
        cCond = "SUM( CASE  WHEN serie IN ('B','CCC') THEN IMPORTE END ) CREDITO, SUM( CASE SERIE WHEN 'Y1' THEN IMPORTE END ) CONTADO "
        Tabla = "FACVENTA_DET"
        crpt = "\vtacoant.rpt"
        cencab = "BODEGA MIGUEL CABRERA"
    ElseIf OptSerie(1).Value = True Then
        cCond = "CREDITO = 0, SUM( CASE  WHEN serie IN ('D2','I2','J2') THEN IMPORTE END ) CONTADO "
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(1).Caption, 1, 2))
        Tabla = IIf(lmismatda, "FACVENTA_DET", "FACVTAMAY")
        crpt = "\vtacantC.rpt"
        cencab = "BODEGA CENTRAL DE ABASTOS"
    ElseIf OptSerie(2).Value = True Then
        cCond = "SUM( CASE  WHEN serie IN ('H2','DDD') THEN IMPORTE END ) CREDITO, SUM( CASE SERIE WHEN 'G2' THEN IMPORTE END ) CONTADO "
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(2).Caption, 1, 2))
        Tabla = IIf(lmismatda, "FACVENTA_DET", "FACVTAMAY")
        crpt = "\vtacoant.rpt"
        cencab = "BODEGA PUERTO ESCONDIDO"
    ElseIf OptSerie(3).Value = True Then
        cCond = "SUM( CASE  WHEN serie IN ('HHH') THEN IMPORTE END ) CREDITO, SUM( CASE  WHEN serie IN ('GGG','B') THEN IMPORTE END ) CONTADO "
        Tabla = "FACVTAMAY"
        crpt = "\vtacoant.rpt"
        cencab = "BODEGA MIAHUATLAN"
    ElseIf OptSerie(5).Value = True Then
        cCond = "SUM( CASE  WHEN serie IN ('KKK','LLL') THEN IMPORTE END ) CREDITO, SUM( CASE  WHEN serie IN ('D','JJJ') THEN IMPORTE END ) CONTADO "
        Tabla = "FACVTAMAY"
        crpt = "\vtacoant.rpt"
        cencab = "BODEGA ISTMO"
    Else: OptSerie(4).Value = True
        cCond = "SUM( CASE WHEN serie IN ('Y1','D2','G2','GGG','I2','D','JJJ') THEN IMPORTE END ) as contado, SUM( case WHEN serie IN ('B','CCC','H2','DDD','J2','HHH','KKK','LLL') THEN IMPORTE END ) as CREDITO "
        Tabla = "FACVENTA_DET"
        crpt = "\vtacoant.rpt"
        cencab = "TODAS LAS BODEGAS DE MAYOREO"
    End If
    rs.Open "SELECT " & cCond & _
                    ", FECHA_DET FROM " & Tabla & " WHERE (FECHA_DET >= '" & DateAdd("YYYY", -1, dtfecini.Value) & "' AND FECHA_DET <= '" & DateAdd("YYYY", -1, dtfecfin.Value) & "') OR " & _
                    "(FECHA_DET >= '" & dtfecini.Value & "' AND FECHA_DET <= '" & dtfecfin.Value & "') " & _
                    "GROUP BY  FECHA_DET ORDER BY FECHA_DET DESC", cn, adOpenStatic, adLockOptimistic, adCmdText
    AñoAnt = Year(rs!fecha_det)
    While Not rs.EOF
       If AñoAnt = Year(rs!fecha_det) Then
          cn.Execute "INSERT INTO vtaconcre(vtacre,vtacon,fecha) VALUES (" & IIf(IsNull(rs!credito), 0, rs!credito) & "," & IIf(IsNull(rs!contado), 0, rs!contado) & ",'" & rs!fecha_det & "')"
       Else
          RST.Open "SELECT * FROM vtaconcre WHERE fecha = '" & DateAdd("YYYY", 1, rs!fecha_det) & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
          If RST.BOF And RST.EOF Then
             cn.Execute "INSERT INTO vtaconcre(vtacreA,vtaconA,fecha) VALUES (" & IIf(IsNull(rs!credito), 0, rs!credito) & "," & IIf(IsNull(rs!contado), 0, rs!contado) & ",'" & DateAdd("yyyy", 1, rs!fecha_det) & "')"
          Else
             cn.Execute "UPDATE vtaconcre SET vtacona = " & IIf(IsNull(rs!contado), 0, rs!contado) & ", vtacrea = " & IIf(IsNull(rs!credito), 0, rs!credito) & " WHERE fecha = '" & DateAdd("yyyy", 1, rs!fecha_det) & "'"
          End If
          RST.Close
       End If
       rs.MoveNext
    Wend
    rs.Close
    rs.Open "SELECT " & cCond & " FROM " & Tabla & " WHERE month(fecha_det) = " & Month(Me.dtfecini.Value) & " And year(fecha_det) = " & Year(DateAdd("yyyy", -1, dtfecini.Value))
    cn.Execute "UPDATE vtaconcre SET avtacon = " & IIf(IsNull(rs!contado), 0, rs!contado) & ",avtacre = " & IIf(IsNull(rs!credito), 0, rs!credito)
    'LAS DEMAS SUCURSALES
    If OptSerie(4).Value = True Then
        Tabla = "FACVTAMAY"
        rs.Close
        rs.Open "SELECT " & cCond & _
                     ", FECHA_DET FROM " & Tabla & " WHERE (FECHA_DET >= '" & DateAdd("YYYY", -1, dtfecini.Value) & "' AND FECHA_DET <= '" & DateAdd("YYYY", -1, dtfecfin.Value) & "') OR " & _
                     "(FECHA_DET >= '" & dtfecini.Value & "' AND FECHA_DET <= '" & dtfecfin.Value & "') " & _
                     "GROUP BY  FECHA_DET ORDER BY FECHA_DET DESC", cn, adOpenStatic, adLockOptimistic, adCmdText
        While Not rs.EOF
           If AñoAnt = Year(rs!fecha_det) Then
              cn.Execute "UPDATE vtaconcre SET vtacon =  vtacon + " & IIf(IsNull(rs!contado), 0, rs!contado) & ", vtacre = vtacre +" & IIf(IsNull(rs!credito), 0, rs!credito) & " WHERE fecha = '" & rs!fecha_det & "'"
           Else
             cn.Execute "UPDATE vtaconcre SET vtacona =  vtacona + " & IIf(IsNull(rs!contado), 0, rs!contado) & ", vtacrea = vtacrea +" & IIf(IsNull(rs!credito), 0, rs!credito) & " WHERE fecha = '" & DateAdd("yyyy", 1, rs!fecha_det) & "'"
           End If
        rs.MoveNext
        Wend
        rs.Close
        rs.Open "SELECT " & cCond & " FROM " & Tabla & " WHERE month(fecha_det) = " & Month(Me.dtfecini.Value) & " And year(fecha_det) = " & Year(DateAdd("yyyy", -1, dtfecini.Value))
        cn.Execute "UPDATE vtaconcre SET avtacon = avtacon + " & IIf(IsNull(rs!contado), 0, rs!contado) & ",avtacre = avtacre +" & IIf(IsNull(rs!credito), 0, rs!credito)
    End If
    cadsql = ""
    Rpt.Formulas(1) = "diat = " & diatra(Month(dtfecfin.Value))
ElseIf OptVentas(4).Value Then
    crpt = "\Despagte.rpt"
    cencab = "VENTAS POR AGENTE DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    If cmbProVta.Text <> "" Then
        cCond = " AND TFPRODUC.CLAPROVE = '" & Trim(Mid(cmbProVta.Text, Len(cmbProVta.Text) - 5)) & "' "
        cencab = cencab + " DE " & Trim(Mid(cmbProVta.Text, 1, Len(cmbProVta.Text) - 5)) & " "
    End If
    cadsql = "SELECT VENTAS.agente, CATCLIENTE.cnombre,FACVENTA_DET.Cantidad, FACVENTA_DET.cantidadp, FACVENTA_DET.importe, FACVENTA_DET.fecha_det, " & _
                   "TFPRODUC.CLAPROVE, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & _
             "FROM PITICO.dbo.VENTAS VENTAS,PITICO.dbo.FACVENTA FACVENTA,PITICO.dbo.CATCLIENTE CATCLIENTE,PITICO.dbo.FACVENTA_DET FACVENTA_DET, PITICO.dbo.TFPRODUC TFPRODUC " & _
             "WHERE VENTAS.noventa = FACVENTA.noventa AND VENTAS.agente *= CATCLIENTE.cclave AND " & _
                   "FACVENTA.numfactura = FACVENTA_DET.factura AND " & _
                   "FACVENTA.serie = FACVENTA_DET.serie AND FACVENTA_DET.Producto = TFPRODUC.CONSEC AND " & _
                   "FACVENTA_DET.importe > 0. AND FACVENTA_DET.fecha_det >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND " & _
                   "FACVENTA_DET.fecha_det <= '" & Format(dtfecfin.Value, "yyyy-dd-mm") & "'" & cCond
ElseIf OptVentas(5).Value Then
   cn.Execute "PARETO '" & dtfecini.Value & "','" & dtfecfin.Value & "'"
   cencab = "LEY DE PARETO"
   crpt = "\Pareto.rpt"
ElseIf OptVentas(7).Value Then
    cencab = "VENTAS FACTURADAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    crpt = "\DESPLA1.rpt"
    If cmbProVta.Text <> "" Then
        cCond = " AND TFPRODUC.CLAPROVE = '" & Trim(Mid(cmbProVta.Text, Len(cmbProVta.Text) - 5)) & "' "
        cencab = cencab + " DE " & Trim(Mid(cmbProVta.Text, 1, Len(cmbProVta.Text) - 5)) & " "
    End If

    cadsql = "SELECT FACVENTA_DET.Cantidad, FACVENTA_DET.cantidadp, FACVENTA_DET.importe, FACVENTA_DET.fecha_det, " & _
                     "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & _
             "FROM PITICO.dbo.FACVENTA_DET FACVENTA_DET, pitico.dbo.TFPRODUC TFPRODUC " & _
             "WHERE FACVENTA_DET.Producto = TFPRODUC.CONSEC AND FACVENTA_DET.importe > 0  AND FACVENTA_DET.fecha_det >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND FACVENTA_DET.fecha_det <= '" & Format(dtfecfin.Value, "yyyy-dd-mm") & "' " & cCond
ElseIf OptVentas(6).Value Then
   If OptSerie(0).Value = True Then
        SUCUR = "BODEGA MIGUEL CABRERA"
        Tabla = "facventa_det"
        cCond = "SUM( CASE serie WHEN 'Y1' THEN importe END) AS CONTADO, SUM( CASE  WHEN serie IN ('B','CCC') THEN importe END) AS Credito"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(0).Caption, 1, 2))
    ElseIf OptSerie(1).Value = True Then
        SUCUR = "BODEGA CENTRAL DE ABASTOS"
        cCond = "SUM( CASE serie WHEN 'D2' THEN importe END) AS CONTADO, Cred = 0 "
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(1).Caption, 1, 2))
    ElseIf OptSerie(2).Value = True Then
        SUCUR = "SUCURSAL PUERTO ESCONDIDO"
        cCond = "SUM( CASE serie WHEN 'G2' THEN importe END) AS CONTADO, SUM( CASE  WHEN serie IN ('H2','DDD') THEN importe END) AS Credito"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(2).Caption, 1, 2))
    ElseIf OptSerie(3).Value = True Then
        SUCUR = "BODEGA MIAHUATLAN"
        cCond = "SUM( CASE serie WHEN 'GGG' THEN importe END) AS CONTADO, SUM( CASE  WHEN serie IN ('HHH') THEN importe END) AS Credito"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(3).Caption, 1, 2))
    ElseIf OptSerie(5).Value = True Then
        SUCUR = "SUCURSAL ISTMO"
        cCond = "SUM( CASE  WHEN serie IN ('JJJ','D') THEN importe END) AS CONTADO, SUM( CASE  WHEN serie IN ('KKK','LLL') THEN importe END) AS Credito"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(5).Caption, 1, 2))
    ElseIf OptSerie(4).Value = True Then
        SUCUR = "TODAS LAS BODEGAS DE MAYOREO"
        cCond = "SUM( CASE  WHEN serie IN ('Y1','D2','G2','GGG','JJJ','D') THEN importe END) AS CONTADO, SUM( CASE  WHEN serie IN ('B','CCC','H2','DDD','HHH','KKK','LLL') THEN importe END) AS Credito"
        lmismatda = (Mid(cSucursal, 1, 2) = Mid(OptSerie(4).Caption, 1, 2))
    End If
    Tabla = IIf(lmismatda, " facventa_det", "facvtamay")
    If OptSerie(4).Value = True Then        'Todas las bodegas
       cn.Execute "DELETE FROM vtaconcre INSERT INTO vtaconcre (cortey1,vtacon,vtacre) SELECT MONTH(fecha_det), " & cCond & " FROM " & Tabla & " WHERE fecha_det >= '" & Me.dtfecini.Value & "' AND FECHA_DET <= '" & Me.dtfecfin.Value & "' GROUP BY MONTH(fecha_det)"
       Set rs = New ADODB.Recordset
       rs.Open "SELECT MONTH(fecha_det) mes, " & cCond & " FROM facventa_det WHERE fecha_det >= '" & Me.dtfecini.Value & "' AND FECHA_DET <= '" & Me.dtfecfin.Value & "' GROUP BY MONTH(fecha_det)", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
       cn.Execute "UPDATE vtaconcre SET vtacon = 0 WHERE vtacon IS NULL "
       cn.Execute "UPDATE vtaconcre SET vtacre = 0 WHERE vtacre IS NULL "
       While Not rs.EOF
         cn.Execute "UPDATE vtaconcre SET vtacon =vtacon + " & rs!contado & ", vtacre = vtacre + " & rs!credito & " WHERE cortey1 = " & rs!mes
         rs.MoveNext
       Wend
       rs.Close
    Else
       cn.Execute "DELETE FROM vtaconcre INSERT INTO vtaconcre (cortey1,vtacon,vtacre) SELECT MONTH(fecha_det), " & cCond & " FROM " & Tabla & " WHERE fecha_det >= '" & Me.dtfecini.Value & "' AND FECHA_DET <= '" & Me.dtfecfin.Value & "' GROUP BY MONTH(fecha_det)"
    End If
    cencab = "FACTOR ESTACIONAL DE VENTAS DEL " & dtfecini.Value & " AL " & dtfecfin.Value
    crpt = "\vtafaest.rpt"
    Rpt.Formulas(1) = "TITULO = '" & SUCUR & "'"
End If
'MsgBox cadsql
Rpt.ReportFileName = App.Path & crpt
Rpt.Connect = cCadConex
Rpt.Formulas(0) = "ENCAB = '" & cencab & "'"
Rpt.WindowTitle = cencab
Rpt.SQLQuery = cadsql
Rpt.Action = 1
stb1.SimpleText = cMensaje
stb1.Refresh
Exit Sub
Error:
   MsgBox Err.Description, vbCritical
   Unload Me
End Sub


Private Sub dtfecfin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub

Private Sub dtfecini_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub

Private Sub Form_Activate()
  Unload frmAreaRecibo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    'SendKeys vbTab
    keybd_event &H9, 0, 0, 0
    'keybd_event VK_TAB, 0, &H2, 0
    KeyAscii = 0
 End If
End Sub

Private Sub Form_Load()
On Error GoTo Error:
'Cargo datos de desplazamiento
 frmAreaRecibo.fraAvance.Visible = True
 frmAreaRecibo.fraAvance.Caption = "Avance de proveedores"
 frmAreaRecibo.fraAvance.Refresh
 frmAreaRecibo.PGB.Min = 0
 cmen = stb1.SimpleText
 stb1.SimpleText = Space(70) + "Espere un momento, cargando catálogo de proveedores"
 stb1.Refresh
 Set rsttemp = New ADODB.Recordset
 rsttemp.CursorType = adOpenDynamic
 rsttemp.Open "SELECT * FROM CATPROV WHERE ACTIVO = 1", cn, adOpenStatic, adLockOptimistic, adCmdText
 frmAreaRecibo.PGB.Max = rsttemp.RecordCount
 rsttemp.Close
 rsttemp.Open "SELECT * FROM CATPROV WHERE ACTIVO = 1", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
 N = 0
 While Not rsttemp.EOF
    N = N + 1
    frmAreaRecibo.PGB.Value = N
    If (Not IsNull(rsttemp!prove)) And Not (IsNull(rsttemp!NOMPROVE)) Then
        cmbProv.AddItem rsttemp!NOMPROVE + "    " + Trim(rsttemp!prove)
        cmbProvPr.AddItem rsttemp!NOMPROVE + "    " + Trim(rsttemp!prove)
        cmbProved.AddItem rsttemp!NOMPROVE + "    " + Trim(rsttemp!prove)
        cmbProvPro.AddItem rsttemp!NOMPROVE + "    " + Trim(rsttemp!prove)
        cmbProVta.AddItem rsttemp!NOMPROVE + "|    " + Trim(rsttemp!prove)
    End If
    rsttemp.MoveNext
 Wend

 'Cargo datos de la pestaÑa de Existencias
 cmbClas(0).AddItem "="
 cmbClas(0).AddItem ">"
 cmbClas(0).AddItem ">="
 cmbClas(0).AddItem "<"
 cmbClas(0).AddItem "<="
 cmbClas(0).ListIndex = 0
 
 frmAreaRecibo.fraAvance.Caption = "Avance de familias"
 frmAreaRecibo.fraAvance.Refresh
 stb1.SimpleText = Space(50) + "Espere un momento, cargando catálogo de familias"
 stb1.Refresh
 rsttemp.Close
 rsttemp.Open "Familias", cCadConex, adOpenStatic, , adCmdTable
 frmAreaRecibo.PGB.Max = rsttemp.RecordCount
 N = 0
 While Not rsttemp.EOF
    N = N + 1
    frmAreaRecibo.PGB.Value = N
    cmbClas(1).AddItem rsttemp!fclave + Space(2) + rsttemp!fdescrip
    rsttemp.MoveNext
 Wend
 
 frmAreaRecibo.fraAvance.Caption = "Avance de lineas"
 frmAreaRecibo.fraAvance.Refresh
 stb1.SimpleText = Space(50) + "Espere un momento, cargando catálogo de lineas"
 stb1.Refresh
 rsttemp.Close
 rsttemp.Open "Select * from Lineas order by SfDescrip ", cCadConex, adOpenStatic, , adCmdText
 frmAreaRecibo.PGB.Max = rsttemp.RecordCount
 rsttemp.Close
 rsttemp.Open "SELECT * FROM Lineas ORDER BY SfDescrip ", cCadConex, adOpenForwardOnly, adLockReadOnly, adCmdText
 N = 0
 While Not rsttemp.EOF
    N = N + 1
    frmAreaRecibo.PGB.Value = N
    cmbClas(2).AddItem rsttemp!sfdescrip + Space(4) + rsttemp!sfclave
    rsttemp.MoveNext
 Wend
 
 frmAreaRecibo.fraAvance.Caption = "Avance de Departamentos"
 frmAreaRecibo.fraAvance.Refresh
 stb1.SimpleText = Space(50) + "Espere un momento, cargando catálogo de Departamentos"
 stb1.Refresh
 rsttemp.Close
 rsttemp.Open "Departamento", cCadConex, adOpenStatic, , Table
 frmAreaRecibo.PGB.Max = rsttemp.RecordCount
 N = 0
 While Not rsttemp.EOF
    N = N + 1
    frmAreaRecibo.PGB.Value = N
    cmbClas(3).AddItem rsttemp!depclave + Space(2) + rsttemp!depdescrip
    rsttemp.MoveNext
 Wend
dtfecini.Value = Format(date, "dd/mm/yyyy")
dtfecfin.Value = Format(date, "dd/mm/yyyy")

cn.Close
cn.ConnectionTimeout = 0
cn.CommandTimeout = 0
cn.Open

'Cal1.Value = Format(date, "dd/mm/yyyy")
stb1.SimpleText = cmen
stb1.Refresh
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmAreaRecibo.Show
End Sub

Private Sub Opprov_Click(Index As Integer)
cmbtipo.Visible = (Index = 7)
End Sub


Private Sub OptAgente_Click(Index As Integer)
cmbCliente.Visible = (Index = 3 Or Index = 6)
cmbAgentes.Visible = (Index <> 3 Or Index <> 6)
End Sub

Private Sub OptExis_Click(Index As Integer)
For N = 0 To 3
    cmbClas(N).Visible = (N = Index)
Next
cmbProved.Visible = False
txtExis.Visible = (Index = 0 Or Index = 6)
cmbClas(0).Visible = (Index = 0 Or Index = 6)
lstNoenv.Visible = Index = 7
If Index = 4 Then
   cmbProved.Visible = True
ElseIf Index = 5 Then
   PORCOSTO = True
End If
End Sub

Private Sub OptPedido_GotFocus(Index As Integer)
txtfecha.Text = date
txtfecha.Visible = (Index = 3) 'Se muestra solo cuando es eficiencia en el surtido de Ped. Aba.
End Sub

Private Sub optPedpro_GotFocus(Index As Integer)
Dim rs As ADODB.Recordset
'   txtPedFecha(0).Visible = (Index = 3 Or Index = 4)
'   txtPedFecha(1).Visible = (Index = 3 Or Index = 4)
'   lblEtiquetas(4).Visible = (Index = 3 Or Index = 4)
'   lblEtiquetas(5).Visible = (Index = 3 Or Index = 4)
'   lblEtiquetas(4).Top = optPedpro(Index).Top
'   lblEtiquetas(5).Top = optPedpro(Index).Top
'   txtPedFecha(0).Top = optPedpro(Index).Top + 250
'   txtPedFecha(1).Top = optPedpro(Index).Top + 250
   lbletiquetas(7).Visible = (Index = 3 Or Index = 4)
   cmbprod.Visible = (Index = 3 Or Index = 4)
   lbletiquetas(7).Caption = IIf(Index = 3, "Producto", "Comprador")
End Sub

Private Sub optPedprove_Click(Index As Integer)
cmbProv.Visible = (Index = 1)
fraPer.Visible = (Index = 1)
'lblEtiquetas.Caption = IIf(Index = 0, "Gerente de compras", "Proveedor")
'txtFecped(0).Visible = (Index = 0)
'txtFecped(1).Visible = (Index = 0)
'lblrtiquetas(0).Visible = (Index = 0)
'lblrtiquetas(1).Visible = (Index = 0)
End Sub

Private Sub OptSerie_Click(Index As Integer)
'Me.FraOrden.Visible = (Index = 4)
'Me.FraTipVta.Visible = (Index = 4)
End Sub

Private Sub OptTrasl_Click(Index As Integer)
  chkPapeleria.Enabled = (Index = 0)
  chkvolumen.Enabled = (Index = 2)
End Sub


Private Sub OptVentas_Click(Index As Integer)
fraSerie.Visible = (Index = 0 Or Index = 3 Or Index = 6)
FraOrden.Visible = fraSerie.Visible And Index = 0
FraTipVta.Visible = fraSerie.Visible And Index = 0
End Sub

Private Sub PEDPROVE()
If Me.optPedprove(0).Value = True Then
   comprador = ""
   'If Me.cmbcompra.Text <> "" Then
   '   comprador = " AND CATPROV.comprador = '" & cmbcompra.Text & "' "
   'End If
   Rpt.Formulas(0) = "ENCABEZADO = 'PEDIDOS RECIBIDOS DEL " & Me.dtfecini.Value & " AL " & dtfecfin.Value & "'"
   Rpt.ReportFileName = App.Path & "\efiprovc.rpt"
   Rpt.WindowTitle = "Eficiencia por comprador en el abasto de mercancia"
   Rpt.SQLQuery = "SELECT DETALLEGLOBAL.dg_cantsol, DETALLEGLOBAL.dg_cantsolp, DETALLEGLOBAL.dg_cantreal, " & _
                           "PEDPROVE.pp_pedido, PEDPROVE.pp_fecrecibe, PEDPROVE.pp_pedBack, " & _
                           "TFPRODUC.CONSEC, " & _
                           "CATPROV.PROVE, CATPROV.NOMPROVE, CATPROV.comprador " & Chr(13) & _
                   "FROM pitico.dbo.DETALLEGLOBAL DETALLEGLOBAL, " & _
                        "pitico.dbo.PEDPROVE PEDPROVE, " & _
                        "pitico.dbo.TFPRODUC TFPRODUC, " & _
                        "pitico.dbo.CATPROV CATPROV " & Chr(13) & _
                   "WHERE  DETALLEGLOBAL.dg_pedido = PEDPROVE.pp_pedido AND " & _
                        "DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND " & _
                        "PEDPROVE.pp_proveedor = CATPROV.PROVE AND " & _
                        "(DETALLEGLOBAL.dg_cantsol > 0 OR DETALLEGLOBAL.dg_cantsolP > 0) AND " & _
                        "PEDPROVE.pp_pedback IS NULL AND PEDPROVE.pp_fecrecibe >= '" & Format(dtfecini.Value, "yyyy-dd-mm") & "' AND PEDPROVE.pp_fecrecibe <= '" & Format(dtfecfin.Value, "yyyy-dd-mm") & "' " & comprador & Chr(13) & _
                   "ORDER BY CATPROV.comprador ASC, CATPROV.NOMPROVE ASC, PEDPROVE.pp_fecrecibe ASC"
Else
   If Trim(cmbProvPr.Text) = "" Then
      MsgBox "ES NECESARIO ESPECIFICAR UN PROVEEDOR", vbInformation
      Exit Sub
   End If
   cMens = Me.stb1.SimpleText
   stb1.SimpleText = Space(55) & "Espere un momento generando reporte.......... "
   stb1.Refresh
   Rpt.WindowTitle = "Llegadas de " & cmbProv.Text
   Rpt.ReportFileName = App.Path & "\llegacdcT.rpt"
   'N = InStr(1, Cmbprov.Text, "[")
   cveprov = Trim(Mid(cmbProvPr.Text, Len(cmbProvPr.Text) - 5))
   'cn.Execute "DELETE FROM llegacdc"
   'cn.Execute "INSERT INTO llegacdc(consec,cajas,dia,mes,año,fecha,back) SELECT CONSEC, SUM(DG_CANTREAL), day(pp_fecrecibe),MONTH(pp_fecrecibe),YEAR(pp_fecrecibe), max(pp_fecrecibe)  FROM PEDPROVE,DETALLEGLOBAL,TFPRODUC WHERE PP_PEDIDO = DG_PEDIDO AND DG_PRODUCTO = CONSEC AND CLAPROVE = '" & cveprov & "' AND NOT DG_CANTREAL IS NULL GROUP BY CONSEC, day(pp_fecrecibe),month( pp_fecrecibe),YEAR(pp_fecrecibe)"
   Rpt.SQLQuery = "SELECT PEDPROVE.pp_fechagen, PEDPROVE.pp_recibe, PEDPROVE.pp_fecrecibe, PEDPROVE.pp_pedback, " & _
                           "DETALLEGLOBAL.dg_pedido, DETALLEGLOBAL.dg_cantsol, DETALLEGLOBAL.dg_cantreal, " & _
                           "TFPRODUC.CLAPROVE, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & Chr(13) & _
                   "FROM pitico.dbo.PEDPROVE PEDPROVE, " & _
                        "pitico.dbo.DETALLEGLOBAL DETALLEGLOBAL, " & _
                        "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                   "WHERE PEDPROVE.pp_pedido = DETALLEGLOBAL.dg_pedido AND " & _
                        "DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND " & _
                        "PEDPROVE.pp_recibe = 1 AND DETALLEGLOBAL.dg_cantsol > 0 AND " & _
                        "TFPRODUC.CLAPROVE = '" & cveprov & "' AND PEDPROVE.pp_pedback IS NULL"
   Rpt.Formulas(0) = "ENCAB = '" & "PRODUCTOS SOLICITADOS DE " & Me.cmbProvPr.Text & "'"
   If optper(0).Value = True Then
      Rpt.SectionFormat(0) = "DETAIL;T;F;F;X;X;X;X"
      Rpt.SectionFormat(1) = "GF3;T;F;F;X;X;X;X"
      Rpt.SectionFormat(2) = "GH3;T;F;F;X;X;X;X"
   ElseIf optper(1).Value = True Then
      Rpt.SectionFormat(0) = "DETAIL;F;F;F;X;X;X;X"
      Rpt.SectionFormat(1) = "GF3;F;F;F;X;X;X;X"
      Rpt.SectionFormat(2) = "GH3;T;F;F;X;X;X;X"
   ElseIf optper(2).Value = True Then
      Rpt.SectionFormat(0) = "DETAIL;F;F;F;X;X;X;X"
      Rpt.SectionFormat(1) = "GF3;F;F;F;X;X;X;X"
      Rpt.SectionFormat(2) = "GH3;F;F;F;X;X;X;X"
   End If
   stb1.SimpleText = cMens
   stb1.Refresh
End If
'MsgBox cRpt.SQLQuery
Rpt.Action = 1
End Sub

Private Sub catprov()
On Error GoTo Error:
Rpt.SQLQuery = ""
If Opprov(0).Value = True Then
    crpt = "\provtipo.rpt"
    condrpt = "{CATPROV.prove} <> '000'"
    Enca = "Reporte General de proveedores"
ElseIf Opprov(1).Value = True Then
    crpt = "\provtipo.rpt"
    condrpt = "{CATPROV.procedencia} = 1"
    Enca = "Reporte de proveedores locales"
ElseIf Opprov(2).Value = True Then
    crpt = "\provtipo.rpt"
    condrpt = "{CATPROV.procedencia} = 0"
    Enca = "Reporte de proveedores foráneos"
ElseIf Opprov(3).Value = True Then
    crpt = "\provcomp.rpt"
    condrpt = "{CATPROV.ACTIVO} = 1"
    Enca = "Listado de proveedores por comprador (a)"
ElseIf Opprov(4).Value = True Then
    crpt = "\provtipo.rpt"
    condrpt = "{CATPROV.ACTIVO} = 1"
    Rpt.Formulas(0) = "TITULO = 'LISTADO DE PROVEEDORES ACTIVOS'"
ElseIf Opprov(5).Value = True Then
    crpt = "\provtipo.rpt"
    condrpt = "{CATPROV.ACTIVO} = 0"
    Enca = "Reporte de proveedores Inactivos"
ElseIf Opprov(6).Value = True Then
    crpt = "\provactin.rpt"
    condrpt = "{CATPROV.PROVE} <>'000'"
    Enca = "Reporte de proveedores con representantes"
ElseIf Opprov(7).Value = True Then
    crpt = "\provtipo.rpt"
    condrpt = "{CATPROV.prove} <> '000'"
ElseIf Opprov(8).Value = True Then
    crpt = "\vtaprov.rpt"
    Rpt.GroupSortFields(0) = "-SUM ({FACVENTA_DET.importe},{TFPRODUC.CLAPROVE})"
    Enca = "Contribución de ventas por proveedor"
    condrpt = "{FACVENTA_DET.fecha_det} >= DATE(2003,10,01) AND {FACVENTA_DET.fecha_det} <= DATE(2003,10,06) and {FACVENTA_DET.rfc_det} <> 'CANC999999999'"
    Rpt.SQLQuery = "SELECT FACVENTA_DET.Cantidad, FACVENTA_DET.cantidadp, FACVENTA_DET.costo, FACVENTA_DET.costop, FACVENTA_DET.importe, FACVENTA_DET.fecha_det, FACVENTA_DET.rfc_det, " & _
                           "TFPRODUC.CLAPROVE, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC, " & _
                           "CATPROV.NOMPROVE " & Chr(13) & _
                   "FROM PITICO.dbo.FACVENTA_DET FACVENTA_DET, pitico.dbo.TFPRODUC TFPRODUC, pitico.dbo.CATPROV CATPROV " & Chr(13) & _
                   "Where FACVENTA_DET.Producto = TFPRODUC.CONSEC AND TFPRODUC.CLAPROVE = CATPROV.PROVE AND " & _
                   "FACVENTA_DET.fecha_det >= '" & Format(Me.dtfecini.Value, "yyyy-dd-mm") & "' AND FACVENTA_DET.fecha_det <= '" & Format(Me.dtfecfin.Value, "yyyy-dd-mm") & _
                   "' and FACVENTA_DET.rfc_det <> 'CANC999999999' " & Chr(13) & _
                   "Order By TFPRODUC.CLAPROVE ASC, FACVENTA_DET.importe ASC"
    
End If
If Opprov(7).Value Then
   Enca = "Reporte General de proveedores"
   If Trim(cmbtipo.Text) <> "" Then
      condrpt = "{CATPROV.tipo} = '" & Mid(cmbtipo.Text, 3, 1) & "'"
      Enca = Enca & " " & cmbtipo.Text
   End If
End If
Rpt.ReportFileName = App.Path & crpt
'Rpt.Formulas(0) = "FORMSELEC = " & condrpt
Rpt.Formulas(1) = "ENCAB = '" & UCase(Enca) & "'"
Rpt.WindowTitle = Enca
'MsgBox Rpt.SQLQuery
Me.Rpt.Action = 1
Exit Sub
Error:
    MsgBox Err.Description
    Unload Me
End Sub


