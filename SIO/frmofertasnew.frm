VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Fofertasnew 
   BackColor       =   &H8000000A&
   Caption         =   "Ofertas de productos"
   ClientHeight    =   6570
   ClientLeft      =   1095
   ClientTop       =   1470
   ClientWidth     =   9990
   Icon            =   "frmofertasnew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9990
   WindowState     =   2  'Maximized
   Begin VB.Frame frmproductos 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   11760
      TabIndex        =   0
      Top             =   6120
      Width           =   11895
      Begin VB.Frame Frmfiltro 
         Caption         =   "Filtrando Productos"
         Height          =   855
         Left            =   3240
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   6015
         Begin VB.PictureBox ProBar1 
            Height          =   500
            Left            =   0
            ScaleHeight     =   435
            ScaleWidth      =   5475
            TabIndex        =   11
            Top             =   0
            Width           =   5535
         End
      End
      Begin VB.CommandButton Cmdsel 
         Height          =   495
         Index           =   3
         Left            =   7920
         Picture         =   "frmofertasnew.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Quitar Todos los renglones "
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Cmdsel 
         Height          =   495
         Index           =   2
         Left            =   6840
         Picture         =   "frmofertasnew.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Quitar Seleccion"
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Cmdsel 
         Height          =   495
         Index           =   1
         Left            =   5040
         Picture         =   "frmofertasnew.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Confirmar Seleccion"
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Cmdsel 
         Height          =   495
         Index           =   0
         Left            =   3960
         Picture         =   "frmofertasnew.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Confirmar Todos "
         Top             =   2520
         Width           =   615
      End
      Begin VB.ListBox Lstoferta 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   4440
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   2760
         Width           =   11415
      End
      Begin VB.ListBox Lstprod 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         ItemData        =   "frmofertasnew.frx":1412
         Left            =   4560
         List            =   "frmofertasnew.frx":1414
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   960
         Width           =   11415
      End
      Begin VB.Label LblOfer 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   6480
         TabIndex        =   16
         Top             =   3840
         Width           =   4935
      End
      Begin VB.Label Lblprod 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   5175
      End
      Begin VB.Label Label5 
         Caption         =   "Productos Confirmados a Ofertar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Productos Seleccionados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   3255
      End
   End
   Begin MSAdodcLib.Adodc Adocompara 
      Height          =   330
      Left            =   8880
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
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
      Caption         =   "Adodc1"
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
   Begin ComctlLib.StatusBar statusbar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   33
      Top             =   6240
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "                                                                                          Para salir presione la tecla [ Esc ]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frmactua 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   11655
      Begin MSMask.MaskEdBox txtfecfin 
         Height          =   315
         Left            =   4320
         TabIndex        =   18
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99/99/99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfecini 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99/99/99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "Ofertas para el área de Mayoreo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   32
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha final"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame frmprecios 
      Height          =   7215
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   11655
      Begin VB.CommandButton cmdofertar 
         Caption         =   "&Quitar Prod."
         Height          =   500
         Index           =   7
         Left            =   1920
         Picture         =   "frmofertasnew.frx":1416
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Desoferta el producto seleccionado"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.PictureBox CR1 
         Height          =   480
         Left            =   10080
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   35
         Top             =   6120
         Width           =   1200
      End
      Begin VB.CommandButton cmdofertar 
         Caption         =   "Actuali&zar"
         Height          =   500
         Index           =   6
         Left            =   7440
         Picture         =   "frmofertasnew.frx":1588
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Actualiza información"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.CommandButton cmdofertar 
         Caption         =   "&Reporte"
         Height          =   500
         Index           =   5
         Left            =   8880
         Picture         =   "frmofertasnew.frx":168A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Reporte de ofertas"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.CommandButton cmdofertar 
         Caption         =   "Regresa&r"
         Height          =   500
         Index           =   4
         Left            =   10320
         Picture         =   "frmofertasnew.frx":1BBC
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.CommandButton cmdofertar 
         Caption         =   "A&plicar"
         Height          =   500
         Index           =   3
         Left            =   4680
         Picture         =   "frmofertasnew.frx":1D2E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Aplica ofertas"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.CommandButton cmdofertar 
         Caption         =   "&Modificar"
         Height          =   500
         Index           =   2
         Left            =   6120
         Picture         =   "frmofertasnew.frx":1E70
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Modifica precios del producto seleccionado"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.CommandButton cmdofertar 
         Caption         =   "&Quitar Tod."
         Height          =   500
         Index           =   1
         Left            =   3360
         Picture         =   "frmofertasnew.frx":1FE2
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Desoferta todo los productos"
         Top             =   6600
         Width           =   1000
      End
      Begin VB.CommandButton cmdofertar 
         Caption         =   "&Agregar"
         Height          =   500
         Index           =   0
         Left            =   600
         Picture         =   "frmofertasnew.frx":2154
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Oferta un nuevo producto"
         Top             =   6600
         Width           =   1000
      End
      Begin MSDataGridLib.DataGrid dgaplica 
         Bindings        =   "frmofertasnew.frx":224E
         Height          =   6375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   11245
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   17
         TabAction       =   2
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
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "Descripcion"
            Caption         =   "                            Producto"
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
            DataField       =   "medida"
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
         BeginProperty Column02 
            DataField       =   "precio1"
            Caption         =   "Pieza Autoserv."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "precio5"
            Caption         =   "1/2 Caj. Envío"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "precio6"
            Caption         =   "1/2 Caj.  Bodega"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Precio2"
            Caption         =   "May.envío y/o crédito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "precio3"
            Caption         =   "    May.  Intermedio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "precio4"
            Caption         =   "    May.   Bodega"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "fechaini"
            Caption         =   "Fecha Inicio"
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
            DataField       =   "producto"
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
         BeginProperty Column10 
            DataField       =   "precio1ant"
            Caption         =   "Autoserv. Normal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "precio5ant"
            Caption         =   "1/2 Caja Env.Normal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "precio6ant"
            Caption         =   "1/2 Caja Bod.Normal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "precio2ant"
            Caption         =   "May.Envío Normal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "precio3ant"
            Caption         =   "May.Interm. Normal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "precio4Ant"
            Caption         =   "May.Bodega Normal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   4889.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column12 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column13 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column14 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column15 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adobus 
         Height          =   735
         Left            =   120
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1296
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
         Caption         =   "Adodc1"
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
   Begin VB.Frame frabusqueda 
      Caption         =   "Búsqueda de Productos:"
      Enabled         =   0   'False
      Height          =   7215
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid gridbusca 
         Bindings        =   "frmofertasnew.frx":2267
         Height          =   5775
         Left            =   120
         Negotiate       =   -1  'True
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   10186
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   0   'False
         HeadLines       =   1.5
         RowHeight       =   15
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "consec"
            Caption         =   "Clave"
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
            DataField       =   "descripc"
            Caption         =   "Descripcion"
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
            DataField       =   "presentacion"
            Caption         =   "Presentación"
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
            DataField       =   "barraspza"
            Caption         =   "Barras"
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
            DataField       =   "ofertado"
            Caption         =   "Ofertado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fecact"
            Caption         =   "Fecha Act"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yy"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   4905.071
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdVerReg 
         Caption         =   "Regresar"
         Height          =   375
         Left            =   9960
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtbusca 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   9375
      End
      Begin VB.Label producto 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   6600
         Width           =   11295
      End
   End
End
Attribute VB_Name = "Fofertasnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fecha1 As Date
Dim fecha2 As Date
Dim lfecha As Boolean
Dim solofechas As Boolean

Private Sub adobus_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo Error:
producto.Caption = adobus.Recordset!descripc & " " & adobus.Recordset!Presentacion
producto.Refresh
Exit Sub
Error:
MsgBox "No se encontro descripcion"
End Sub

Private Sub cmdsalir_Click()
On Error GoTo Error:
Unload Me
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdofertar_Click(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0
    'mostrar la pantalla de productos
    frabusqueda.Enabled = True
    frabusqueda.Visible = True
    frmprecios.Visible = False
    txtbusca.SetFocus
Case 1
    If MsgBox("REALMENTE DESEAS QUITAR LAS OFERTAS DEL " & txtfecini.Text & " AL " & txtfecfin.Text, vbQuestion + vbYesNo, "Ofertas") = vbYes Then
       cn.Execute "UPDATE preprod SET precio1 = t.precio1ant, precio2 = t.precio2ant, precio3 = t.precio3ant, precio4 = t.precio4ant, precio5 = t.precio5ant, precio6 = t.precio6ant FROM preciostemp t WHERE producto = preclave AND fechaini = '" & txtfecini.Text & "' AND fechafin = '" & txtfecfin.Text & "'"
       cn.Execute "UPDATE tfproduc SET ofertado = 0, ACTUALIZADO = 1, FECACT = '" & date & "' FROM preciostemp t WHERE producto = CONSEC AND T.fechaini = '" & txtfecini.Text & "' AND T.fechafin = '" & txtfecfin.Text & "'"
       cn.Execute "DELETE FROM preciostemp WHERE fechaini = '" & txtfecini.Text & "' AND fechafin = '" & txtfecfin.Text & "'"
       Adocompara.Refresh
    End If
Case 7
   'ES MOMENTO DE QUITAR UNA OFERTA
   Call quitaoferta
Case 2
     'MsgBox "UPDATE preprod SET precio1 = t.precio1, precio2 = t.precio2, precio3 = t.precio3, precio4 = t.precio4, precio5 = t.precio5, precio6 = t.precio6, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave and producto = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     'cn.Execute "UPDATE preprod SET precio1 = t.precio1, precio2 = t.precio2, precio3 = t.precio3, precio4 = t.precio4, precio5 = t.precio5, precio6 = t.precio6, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave and producto = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     cn.Execute "UPDATE preprod SET precio1 = t.precio1, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio1 > 0 and producto = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     cn.Execute "UPDATE preprod SET precio2 = t.precio2, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio2 > 0 and producto = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     cn.Execute "UPDATE preprod SET precio3 = t.precio3, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio3 > 0 and producto = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     cn.Execute "UPDATE preprod SET precio4 = t.precio4, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio4 > 0 and producto = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     cn.Execute "UPDATE preprod SET precio5 = t.precio5, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio5 > 0 and producto = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     cn.Execute "UPDATE preprod SET precio6 = t.precio6, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio6 > 0 and producto = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     cn.Execute "UPDATE TFPRODUC SET ofertado = 1, usuario = '" & cUsuario & "', fecact = '" & date & "' FROM preciostemp t WHERE producto = consec and consec = '" & Trim(Me.Adocompara.Recordset!producto) & "'"
     MsgBox "EL PRODUCTO SE ACTUALIZO CORRECTAMENTE", vbInformation
Case 3
    If MsgBox("CONFIRMA SI DESEAS APLICAR OFERTAS?", vbQuestion + vbYesNo) = vbYes Then
        'cn.Execute "UPDATE preprod SET precio1 = t.precio1, precio2 = t.precio2, precio3 = t.precio3, precio4 = t.precio4, precio5 = t.precio5, precio6 = t.precio6, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave"
        cn.Execute "UPDATE preprod SET precio1 = t.precio1, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio1 > 0"
        cn.Execute "UPDATE preprod SET precio2 = t.precio2, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio2 > 0"
        cn.Execute "UPDATE preprod SET precio3 = t.precio3, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio3 > 0"
        cn.Execute "UPDATE preprod SET precio4 = t.precio4, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio4 > 0"
        cn.Execute "UPDATE preprod SET precio5 = t.precio5, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio5 > 0"
        cn.Execute "UPDATE preprod SET precio6 = t.precio6, prusuario = '" & cUsuario & "', fechaact = '" & date + Time & "' FROM preciostemp t WHERE producto = preclave AND t.precio6 > 0"
        
        cn.Execute "UPDATE TFPRODUC SET ofertado = 1, usuario = '" & cUsuario & "', fecact = '" & date & "' FROM preciostemp t WHERE producto = consec"
        cn.Execute "UPDATE preciostemp SET activado = 1"
        MsgBox "LOS CAMBIOS DE PRECIOS PARA OFERTAS SE ACTUALIZARON CORRECTAMENTE", vbInformation
    End If
Case 4
    Unload Me
Case 5
    cr1.WindowTitle = "Listado de productos ofertados"
    cr1.ReportFileName = App.Path & "\ofertas.rpt"
    cr1.Connect = strconnect
    cr1.Action = 1
Case 6
   Adocompara.Refresh
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdVerReg_Click()
 Me.frabusqueda.Visible = False
 frmprecios.Visible = True
End Sub

Private Sub dgaplica_AfterColUpdate(ByVal ColIndex As Integer)
If Adocompara.Recordset!PRECIO3 > Adocompara.Recordset!PRECIO2 Then Adocompara.Recordset!PRECIO3 = Adocompara.Recordset!PRECIO2

If UCase(Me.dgaplica.Columns(ColIndex).DataField) = "PRECIO4" Then
   SendKeys "{DOWN}"
End If
End Sub

Private Sub dgaplica_BeforeUpdate(Cancel As Integer)
Dim RST As ADODB.Recordset
Set RST = New ADODB.Recordset
RST.Open "SELECT * FROM tfproduc WHERE CONSEC =  " & Trim(Adocompara.Recordset!producto), cn, adOpenKeyset, adLockOptimistic, adCmdText
'If Round(RST!PRECOSTO, 2) > Round(Adocompara.Recordset!precio4, 2) Or Round(RST!PRECOSTO, 2) >= Round(Adocompara.Recordset!PRECIO3, 2) Or Round(RST!PRECOSTO, 2) >= Round(Adocompara.Recordset!PRECIO2, 2) Then
If Round(RST!PRECOSTO, 2) > Round(Adocompara.Recordset!precio4, 2) Or Round(RST!PRECOSTO, 2) >= Round(Adocompara.Recordset!PRECIO3, 2) Or Round(RST!PRECOSTO, 2) >= Round(Adocompara.Recordset!PRECIO2, 2) Then
   MsgBox "EL PRECIO DE COSTO ES MAYOR O IGUAL AL PRECIO DE BODEGA", vbCritical
   Cancel = True
End If
RST.Close
Set RST = Nothing
End Sub

Private Sub dgaplica_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Form_Activate()
Unload frmAreaRecibo
dgaplica.SetFocus
If txtfecini.Text = "__/__/__" Then
    Frmactua.Visible = True
    Frmactua.Enabled = True
    txtfecini.SetFocus
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 117
        'AGREGAR PRODUCTOS
        frabusqueda.Enabled = True
        frabusqueda.Visible = True
        txtbusca.SetFocus
    Case 118
        'QUITAR PRODUCTO
        Call quitaoferta
    Case 119
         frmCalc.Show 1
    Case 120
         'POR PROVEEDOR
    Case 27
        Unload Me
End Select
End Sub

Private Sub Form_Load()
On Error GoTo Error:
solofechas = True
' para poder realizar las consultas se abre un controL de datos
adobus.CursorType = adOpenKeyset
adobus.LockType = adLockOptimistic
adobus.CommandType = adCmdText
adobus.ConnectionString = cCadConex '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
adobus.RecordSource = "SELECT fecact,consec,activo,claprove,descripc,nomcorto,STR(PAQUETES) + ' X ' + STR(CONTENID,8,3)+ ' ' + MEDIDA AS PRESENTACION,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja,OFERTADO" & _
" FROM tfproduc WHERE activo = 1 ORDER BY descripc"
adobus.Refresh
'LA VISTA DE LOS PRODUCTOS OFERTADOS
Adocompara.CommandType = adCmdText
Adocompara.ConnectionString = cCadConex
Adocompara.RecordSource = "SELECT * FROM preciostemp order by descripcion"
Adocompara.Refresh
If Not (Adocompara.Recordset.BOF And Adocompara.Recordset.EOF) Then
   txtfecini.Text = Adocompara.Recordset!FechaIni
   txtfecfin.Text = Adocompara.Recordset!FechaFin
End If
'que pregunte las fechas desde el principio
Frmactua.Enabled = True
Frmactua.Visible = True
'MsgBox Adocompara.Recordset.RecordCount
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmAreaRecibo.Show
End Sub

Private Sub gridbusca_KeyPress(KeyAscii As Integer)
Dim varBmk As Variant
If KeyAscii = 13 Then
'On Error GoTo error:
strcveprov = adobus.Recordset!claprove
'EN ESTA PARTE SE DEBE DE AGREGAR A LA TABLA DE PRECIOS TEMP DIRECTAMENTE PARA OFERTAR
If txtfecini.Text = "__/__/__" Then
    'EN EL CASO DE QUE NO SE HAYAN PUESTO LAS FECHAS
    Frmactua.Enabled = True
    Frmactua.Visible = True
    Me.txtfecini.SetFocus
Else
   Call nuevaoferta(gridbusca.Columns(0).Text)
End If
txtbusca.Text = ""
frmprecios.Enabled = True
frmprecios.Visible = True
Me.dgaplica.SetFocus
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtbusca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtbusca.Text = UCase(Trim(txtbusca.Text))
    Dim consulta As String
    consulta = " descripc like '" & Trim(txtbusca.Text) & "%'"
    adobus.Recordset.MoveFirst
    If consulta <> " descripc like '*'" Then
       adobus.Recordset.Find consulta
    End If
    If adobus.Recordset.EOF Then
        MsgBox "No se encontró ningun producto"
    Else
        gridbusca.Enabled = True
        gridbusca.Visible = True
        gridbusca.SetFocus
    End If
End If

End Sub

Private Sub Txtfecfin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Not solofechas Then
       Call nuevaoferta(gridbusca.Columns(0).Text)
   End If
   Frmactua.Visible = False
End If
End Sub

Private Sub Txtfecfin_LostFocus()
On Error GoTo Error:
moy = CDate(txtfecfin.Text)
If txtfecfin.Text = "  /  /  " Then
   MsgBox "Fecha Invalida..."
End If
Exit Sub
Error:
  MsgBox "Fecha Invalida"
  txtfecini.SetFocus
End Sub

Private Sub Txtfecini_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.txtfecfin.SetFocus
End If
End Sub

Private Sub Txtfecini_LostFocus()
On Error GoTo Error:
moy = CDate(txtfecini.Text)
If txtfecini.Text = "  /  /  " Then
   MsgBox "Fecha Invalida..."
End If
Exit Sub
Error:
  MsgBox "Fecha Invalida", vbInformation
  txtfecini.SetFocus
End Sub

Sub nuevaoferta(clave As String)
'On Error Resume Next
Dim rspre As ADODB.Recordset
Set rspre = New ADODB.Recordset
 If MsgBox("CONFIRMA SI DESEAS OFERTAR EL PRODUCTO" & Chr(13) & gridbusca.Columns(1).Text + " " + gridbusca.Columns(2).Text, vbQuestion + vbYesNo, "Confirmar oferta") = vbNo Then
    Exit Sub
 End If
 rspre.Open "SELECT * FROM preciostemp WHERE producto = '" & Trim(clave) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
 If rspre.EOF Then
    descripcion = Me.gridbusca.Columns(1).Text & "   " & gridbusca.Columns(2).Text
    'PARA MANEJARLO COMO CADENA DE CONEXION, SE DEBE HACER  EN DOS PASOS
    cn.BeginTrans
    CADENA = " INSERT INTO preciostemp ( Producto,Descripcion,medida,precio1ant,precio2ant,precio3ant,precio4ant,precio5ant,precio6ant,precio1,precio2,precio3,precio4,precio5,precio6 ) SELECT consec, descripc, ltrim(str(paquetes)) + ' X ' + ltrim(str(contenid,10,3)) + space(2) + medida,   preprod.precio1 , preprod.precio2 , preprod.precio3 ,preprod.precio4, preprod.precio5, preprod.precio6 " & _
    ", preprod.precio1 , preprod.precio2 , preprod.precio3 , preprod.precio4 , preprod.precio5, preprod.precio6 " & _
    " from preprod,tfproduc where preclave = '" & Trim(clave) & "' and consec = preclave "
    cn.Execute CADENA
    'LA SEGUNDA PARTE ES UN UPDATE PARA LAS FECHAS
    CADENA = "UPDATE preciostemp SET fechaini = '" & Trim(txtfecini.Text) & "', fechafin = '" & Trim(txtfecfin.Text) & "'  WHERE producto = '" & Trim(clave) & "'"
    cn.Execute CADENA
    cn.CommitTrans
    Adocompara.Refresh
    Adocompara.Recordset.MoveFirst
    Adocompara.Recordset.Find "PRODUCTO = '" & Trim(clave) & "'"
Else
   MsgBox "ESTE PRODUCTO SE ENCUENTRA OFERTADO !!!", vbInformation
End If
rspre.Close
Me.frabusqueda.Enabled = False
Me.frabusqueda.Visible = False
Exit Sub
Error:
MsgBox Err.Description
End Sub

Sub quitaoferta()
On Error Resume Next
    respsn = MsgBox("REALMENTE DESEA ELIMINAR LA OFERTA DE " & Me.dgaplica.Columns(0).Text & " " & Me.dgaplica.Columns(1).Text, vbExclamation + vbYesNo, "OFERTAS")
    If respsn = vbNo Then
        Exit Sub
    End If
    cveprod = Me.dgaplica.Columns(9).Text
    CADENA = "UPDATE PREPROD SET precio1 = t.precio1ant, precio2 = t.precio2ant, precio3 = t.precio3ant, precio4 = t.precio4ant, precio5 = t.precio5ant, precio6 = t.precio6ant FROM preciostemp t WHERE producto = preclave and producto = '" & Trim(cveprod) & "'"
    cn.Execute CADENA
    CADENA = "UPDATE TFPRODUC SET  ofertado = 0, ACTUALIZADO = 1, FECACT = '" & date + Time & "' WHERE CONSEC = '" & Trim(cveprod) & "'"
    cn.Execute CADENA
    ' DESPUES DE RECALCULAR EL PRECIO SE DEBE DE BORRAR DE PRECIOSTEMP
    cn.Execute "DELETE FROM preciostemp WHERE producto = " & Trim(cveprod)
    Adocompara.Refresh
    
Exit Sub
Error:
MsgBox Err.Description
End Sub
