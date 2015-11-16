VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrincipal 
   BackColor       =   &H8000000B&
   Caption         =   "Punto de venta mayoreo"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8880
   Icon            =   "frmPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FraPend 
      Height          =   6180
      Left            =   120
      TabIndex        =   40
      Top             =   1560
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton Command1 
         Caption         =   "&Regresar"
         Height          =   525
         Left            =   10800
         Picture         =   "frmPrincipal.frx":400A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Importa pedido de franquicias"
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdborravta 
         Caption         =   "&Borrar"
         Height          =   525
         Left            =   10800
         Picture         =   "frmPrincipal.frx":417C
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Importa pedido de franquicias"
         Top             =   240
         Width           =   855
      End
      Begin MSAdodcLib.Adodc AdoPend 
         Height          =   330
         Left            =   4320
         Top             =   5520
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
         Caption         =   "AdoPendientes"
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
      Begin MSDataGridLib.DataGrid DBGRDPend 
         Bindings        =   "frmPrincipal.frx":42EE
         Height          =   5895
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   14680063
         ForeColor       =   8388608
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "noventa"
            Caption         =   "Fol.Vta."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "folpreventa"
            Caption         =   "Pvta"
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
            DataField       =   "agente"
            Caption         =   "Agente"
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
            DataField       =   "fecha"
            Caption         =   "       Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "descripc"
            Caption         =   "Descripción"
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
            DataField       =   "present"
            Caption         =   "Present"
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
            DataField       =   "cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "precio"
            Caption         =   "Precio"
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
         BeginProperty Column08 
            DataField       =   "importe"
            Caption         =   "importe"
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
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   4289.953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   840.189
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc AdoFranq 
      Height          =   330
      Left            =   2040
      Top             =   0
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
      Caption         =   "AdoVentas"
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
   Begin MSAdodcLib.Adodc Adodbf 
      Height          =   330
      Left            =   600
      Top             =   6120
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
   Begin MSComDlg.CommonDialog cmdg 
      Left            =   720
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraCon 
      BackColor       =   &H00C0C000&
      Caption         =   "Contraseña de acceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      TabIndex        =   22
      Top             =   3480
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
         BackColor       =   &H00C0C000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FraCambia 
      Caption         =   "CAMBIAR SITUACION DE LA VENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3720
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdModReg 
         Caption         =   "Regresar"
         Height          =   375
         Left            =   2760
         TabIndex        =   34
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Optsitua 
         Caption         =   "Venta cobrada (2)"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   32
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Optsitua 
         Caption         =   "Venta facturada credito (3)"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   31
         Top             =   1080
         Width           =   2295
      End
      Begin VB.OptionButton Optsitua 
         Caption         =   "Venta en tramite (1)"
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   30
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.TextBox txtcambia 
         Alignment       =   2  'Center
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
         Left            =   600
         TabIndex        =   29
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdModAce 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblEtiquetas 
         Alignment       =   2  'Center
         Caption         =   "CorteY1"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame FraOpcion 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmdPend 
         Height          =   400
         Left            =   5040
         Picture         =   "frmPrincipal.frx":4304
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Ventas pendientes de facturar"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdtmp 
         Caption         =   "Temporal"
         Height          =   255
         Left            =   6720
         TabIndex        =   39
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   3
         Left            =   3120
         Picture         =   "frmPrincipal.frx":4646
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ir al ultimo"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Enabled         =   0   'False
         Height          =   400
         Index           =   1
         Left            =   6120
         Picture         =   "frmPrincipal.frx":47B8
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Catálogo de clientes"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   0
         Left            =   2640
         Picture         =   "frmPrincipal.frx":48BA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ir al primero"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   2
         Left            =   4320
         Picture         =   "frmPrincipal.frx":4A2C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Actualizar datos para reflejar cambios realizados por otros usuarios"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdImpFra 
         Height          =   400
         Left            =   5640
         Picture         =   "frmPrincipal.frx":4B2E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Importa pedido de franquicias"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   2
         Left            =   1800
         Picture         =   "frmPrincipal.frx":4C70
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ver historial de creditos"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   3
         Left            =   1920
         Picture         =   "frmPrincipal.frx":4D6A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancelacion de una Venta"
         Top             =   240
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Cancel          =   -1  'True
         Height          =   405
         Index           =   4
         Left            =   11040
         Picture         =   "frmPrincipal.frx":4EDC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbFiltro 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":504E
         Left            =   8280
         List            =   "frmPrincipal.frx":5050
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   1
         Left            =   1080
         Picture         =   "frmPrincipal.frx":5052
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cobrar la venta seleccionada"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   5
         Left            =   600
         Picture         =   "frmPrincipal.frx":5154
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Modificar la venta seleccionada actualmente"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   0
         Left            =   120
         Picture         =   "frmPrincipal.frx":52C6
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Genera una nueva venta"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   4
         Left            =   3840
         Picture         =   "frmPrincipal.frx":5408
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Busca folio unico de la venta"
         Top             =   240
         Width           =   500
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de ventas: 999,999"
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
         Left            =   6720
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblFiltro 
         Alignment       =   2  'Center
         Caption         =   "Filtro"
         Height          =   255
         Left            =   8160
         TabIndex        =   19
         Top             =   120
         Width           =   2775
      End
   End
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   10905
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   9525
            MinWidth        =   9525
            Text            =   "Click en el encabezado ordena los datos en base a la columna"
            TextSave        =   "Click en el encabezado ordena los datos en base a la columna"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "F9=>INICIAR VTAS."
            TextSave        =   "F9=>INICIAR VTAS."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2823
            MinWidth        =   2823
            Text            =   "F10=>MOD.STATUS"
            TextSave        =   "F10=>MOD.STATUS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2823
            MinWidth        =   2823
            Text            =   "F11 => Quitar Aum."
            TextSave        =   "F11 => Quitar Aum."
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Quita el incremento por flete"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2824
            MinWidth        =   2824
            Text            =   "F12 => CORTE Z R"
            TextSave        =   "F12 => CORTE Z R"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoVentas 
      Height          =   330
      Left            =   600
      Top             =   6480
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
      Caption         =   "AdoVentas"
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
   Begin MSDataGridLib.DataGrid dbgrdVta 
      Bindings        =   "frmPrincipal.frx":5502
      Height          =   6180
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10901
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   15657960
      ColumnHeaders   =   -1  'True
      ForeColor       =   4194368
      HeadLines       =   1.5
      RowHeight       =   15
      RowDividerStyle =   3
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "Noventa"
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
         DataField       =   "FOLIOINFINITO"
         Caption         =   "Fol.Inf."
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
         DataField       =   "FolioVenta"
         Caption         =   "FolVta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Fecha"
         Caption         =   "         Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "cl_terminal"
         Caption         =   "Modulo"
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
      BeginProperty Column05 
         DataField       =   "cnombre"
         Caption         =   "                         Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "MontoTotal"
         Caption         =   "  Importe"
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
      BeginProperty Column07 
         DataField       =   "situacion"
         Caption         =   "Modo"
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
      BeginProperty Column08 
         DataField       =   "credito"
         Caption         =   "Cred."
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
      BeginProperty Column09 
         DataField       =   "FolPreventa"
         Caption         =   "PVta"
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
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3674.835
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   629.858
         EndProperty
      EndProperty
   End
   Begin VB.Frame fradescripcion 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   11775
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   0
         Left            =   7920
         TabIndex        =   37
         Top             =   330
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   20119555
         CurrentDate     =   36892
      End
      Begin VB.TextBox txtcliente 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5400
         TabIndex        =   20
         Top             =   360
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   1
         Left            =   9720
         TabIndex        =   38
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   20119555
         CurrentDate     =   36892
      End
      Begin VB.Label lblCliente 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Cliente"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   5400
         TabIndex        =   21
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblSituacion 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   320
         Width           =   3495
      End
      Begin VB.Label lblSit 
         BackColor       =   &H80000004&
         Caption         =   "Situacion de la venta"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblEtiquetas 
         Alignment       =   2  'Center
         Caption         =   "Fecha Inicial"
         Height          =   255
         Index           =   0
         Left            =   7920
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblEtiquetas 
         Alignment       =   2  'Center
         Caption         =   "Fecha Final"
         Height          =   255
         Index           =   1
         Left            =   9720
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Menu mnuventas 
      Caption         =   "&Ventas"
      Begin VB.Menu mnunueva 
         Caption         =   "&Nueva"
      End
      Begin VB.Menu mnumodi 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnucobro 
         Caption         =   "&Cobro"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnufacvta 
         Caption         =   "Facturar Venta"
      End
      Begin VB.Menu mnucancela 
         Caption         =   "Ca&ncelar"
      End
      Begin VB.Menu mnuhisto 
         Caption         =   "&Historial"
      End
      Begin VB.Menu mnufecha 
         Caption         =   "&Fecha"
      End
   End
   Begin VB.Menu mnusal 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cCond As String
Private cFecha As String
Private rstCli As ADODB.Recordset
Private ntext As Integer
Private nConPrin As Integer
'Variables para el corte X

Private Sub AdoVentas_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next

If AdoVentas.Recordset!situacion = "1" Then
   lblSituacion.Caption = "VENTA A MAYOREO EN TRAMITE"
ElseIf AdoVentas.Recordset!situacion = "2" Then
   lblSituacion.Caption = "VENTA A MAYOREO COBRADA"
ElseIf AdoVentas.Recordset!situacion = "3" Then
   lblSituacion.Caption = "VENTA A CREDITO NO COBRADA"
Else
   lblSituacion.Caption = ""
End If
cmdopcion(5).Enabled = (AdoVentas.Recordset!situacion <= "1")
mnumodi.Enabled = (AdoVentas.Recordset!situacion <= "1")
End Sub

Private Sub cmbFiltro_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub cmbFiltro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     KeyAscii = 0
     SendKeys "{TAB}"
  End If
End Sub

Private Sub cmbFiltro_LostFocus()
On Error GoTo Error:
Select Case cmbFiltro.ListIndex
Case 0 'Todos
    cCond = " NoVenta > 0 "
    ccondrpt = "{PEDIDOS.P_pedido} <> '' "
Case 1 'En tramite CONTADO
    cCond = " situacion <= '1' AND credito = 0 AND prevta = 0"
    ccondrpt = "{PEDIDOS.P_situacion} <= 1"
Case 2 'En tramite CREDITO
    cCond = " situacion <= '1' AND (credito = 1 OR prevta = 1)"
    ccondrpt = "{PEDIDOS.P_situacion} <= 1"
Case 3 'Cobradas
    cCond = " situacion = '2' "
    ccondrpt = "{PEDIDOS.P_situacion} = 2"
End Select
  
cFecha = " AND fecha >= '" & Format(dtpFecha(0).Value, "DD-MM-YYYY") & "' AND fecha <= '" & Format(DateAdd("d", 1, dtpFecha(1).Value), "DD-MM-YYYY") & "'"        'Cargo todas las ventas
AdoVentas.ConnectionString = cCadConex
AdoVentas.CommandType = adCmdText
AdoVentas.RecordSource = "SELECT * FROM Ventas, CatCliente WHERE clcliente = cClave AND " & cCond & cFecha & " ORDER BY cNombre"
AdoVentas.Refresh
For N = 0 To 4
    Cmdmoverse(N).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
Next
 'cmdopcion(0).Enabled = ModVta
 cmdopcion(1).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
 cmdopcion(5).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
lblInfo.Caption = Str(AdoVentas.Recordset.RecordCount)
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdConAceptar_Click()
Dim RsCon As ADODB.Recordset
    Set RsCon = New ADODB.Recordset
    RsCon.Open "SELECT * FROM USUARIOS WHERE PASSMASTER = '" & Trim(txtContra.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If RsCon.RecordCount = 0 Then
        MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
        txtContra.SetFocus
        SendKeys "+{HOME}"
        Exit Sub
    Else
    If nConPrin = 0 Then
       If MsgBox("REALMENTE DESEAS INICIALIZAR LAS VENTAS" & Chr(13) & "DE LA COMPUTADORA " & Caja, vbQuestion + vbYesNo) = vbYes Then
          cn.Execute "UPDATE folios SET folioventa = 0 WHERE caja = '" & Caja & "'"
          MsgBox "LAS VENTAS SE INICIALIZARON CORRECTAMENTE", vbInformation
       End If
    ElseIf nConPrin = 1 Then
       Me.fracambia.Visible = True
    End If
    fraCon.Visible = False
End If

End Sub

Private Sub cmdConCance_Click()
  Me.fraCon.Visible = False
End Sub

Private Sub CmdImpFra_Click()
Dim cmen As String
Dim lTrans As Boolean
On Error GoTo Error:
    cmdg.DialogTitle = "Importar pedido de franquicias"
    cmdg.Filter = "Pedidos de franquicias ( MAY???.dbf ) | MAY???.dbf"
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
    If (cArch = "MAYTLA") And AdoVentas.Recordset!clcliente <> 4179 Then
       MsgBox "EL CLIENTE CON CLAVE 4179 FRANQUICIA TLACOLULA ESTA ASOCIADO CON EL ARCHIVO MAYTLA, POR LO TANTO NO SE PUEDE IMPORTAR ESTE ARCHIVO", vbCritical
       Exit Sub
    ElseIf (cArch = "MAYZIM") And AdoVentas.Recordset!clcliente <> 12412 Then
       MsgBox "EL CLIENTE CON CLAVE 12412 FRANQUICIA ZIMATLAN ESTA ASOCIADO CON EL ARCHIVO MAYZIM, POR LO TANTO NO SE PUEDE IMPORTAR ESTE ARCHIVO", vbCritical
       Exit Sub
    ElseIf (cArch = "MAYMIA") And AdoVentas.Recordset!clcliente <> 1630 Then
       MsgBox "EL CLIENTE CON CLAVE 1630 FRANQUICIA MIAHUATLAN ESTA ASOCIADO CON EL ARCHIVO MAYMIA, POR LO TANTO NO SE PUEDE IMPORTAR ESTE ARCHIVO", vbCritical
       Exit Sub
    End If

    AdoDbf.CommandType = adCmdText
    AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
    AdoDbf.RecordSource = "SELECT * FROM " & cArch & " WHERE ventac = 0 AND ventap = 0"
    AdoDbf.Refresh
    If AdoDbf.Recordset.RecordCount > 0 Then
       MsgBox "A CONTINUACION SE PROCESARAN " & AdoDbf.Recordset.RecordCount & " PRODUCTOS ", vbInformation, "Importa pedido de franquicias"
    Else
       MsgBox "EN EL PEDIDO DE MAYOREO NO EXISTEN PRODUCTOS PARA DAR DE ALTA", vbInformation, "Franquicias"
       Exit Sub
    End If
    'Se agrega el detalle de la venta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    cmen = Stb1.Panels(1).Text
    lTrans = True
    cn.BeginTrans
    While Not AdoDbf.Recordset.EOF
        Stb1.Panels(1).Text = "Procesando: " & Trim(AdoDbf.Recordset!descripc) & " " & AdoDbf.Recordset!medida
        Stb1.Refresh
        rs.Open "SELECT * FROM inventario WHERE inprod = " & AdoDbf.Recordset!CONSEC, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
        'Cuando piden solo cajas
        If AdoDbf.Recordset!pedirc > 0 And AdoDbf.Recordset!pedirp = 0 And AdoDbf.Recordset!VENTAC = 0 Then
           If rs!InCant >= AdoDbf.Recordset!pedirc Then
              AdoDbf.Recordset!VENTAC = AdoDbf.Recordset!pedirc
              AdoDbf.Recordset.Update
              cn.Execute "INSERT INTO ventas_det(noventa,cl_producto,cantidad) VALUES (" & AdoVentas.Recordset!noventa & ",'" & AdoDbf.Recordset!CONSEC & "'," & AdoDbf.Recordset!pedirc & ")"
              cn.Execute "UPDATE inventario SET incant = incant - " & AdoDbf.Recordset!pedirc & " WHERE inprod = '" & AdoDbf.Recordset!CONSEC & "'"
           ElseIf rs!InCant <= AdoDbf.Recordset!pedirc And rs!InCant > 0 And AdoDbf.Recordset!VENTAC = 0 Then
              AdoDbf.Recordset!VENTAC = rs!InCant
              AdoDbf.Recordset.Update
              cn.Execute "INSERT INTO ventas_det(noventa,cl_producto,cantidad) VALUES (" & AdoVentas.Recordset!noventa & ",'" & AdoDbf.Recordset!CONSEC & "'," & rs!InCant & ")"
              cn.Execute "UPDATE inventario SET incant = incant - " & rs!InCant & " WHERE inprod = '" & AdoDbf.Recordset!CONSEC & "'"
           End If
        End If
        rs.Close
        AdoDbf.Recordset.MoveNext
    Wend
    cn.Execute "UPDATE ventas_det SET precio = precio4, preciop = precio1, ventas_det.precosto = t.precosto, ventas_det.precostop = t.precosto / t.paquetes, importe = (precio4 * cantidad) + (precio1 * cantidadp)," & _
                      "ventas_det.ieps =t.ieps,ventas_det.iva= t.iva,ventas_det.tasaieps = t.tasaieps FROM tfproduc T, preprod P WHERE preclave = consec AND consec = cl_producto AND preclave = cl_producto AND noventa = " & AdoVentas.Recordset!noventa
    cn.Execute "UPDATE ventas_det SET ventas_det.iva = c.iva, ventas_det.ieps = c.ieps FROM cargos C WHERE caprod = cl_producto AND noventa = " & AdoVentas.Recordset!noventa

    cn.CommitTrans: lTrans = False
    rs.Open "SELECT SUM(Importe) AS Subto FROM VENTAS_DET, TFPRODUC WHERE consec = cl_producto AND NOVENTA = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
    'ACTUALIZANDO EL MONTO DE LA VENTA EN GENERAL
    cn.Execute "UPDATE ventas SET MontoTotal = " & rs!subto & " WHERE Noventa = " & AdoVentas.Recordset!noventa
    rs.Close

    Stb1.Panels(1).Text = "Generando reporte de faltantes...."
    Stb1.Refresh
    Open App.Path & "\FAL" & Mid(cArch, 4) & ".TXT" For Output As #1 ' Abre el archivo para operaciones de salida.
    Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
    Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
    Print #1, "FALTANTES DEL PEDIDO DE FRANQUICIA "; cArch; "  "; UCase(Format(date, "long date"))
    Print #1,   ' Imprime una línea en blanco en el archivo.
    Print #1, "========================================================================================="
    Print #1, "    DESCRIPCION                      MEDIDA                SOLC SURC DIFC  SOLC SURC DIFC"
    Print #1, "========================================================================================="
    AdoDbf.RecordSource = "SELECT * FROM " & cArch
    AdoDbf.Refresh
    Dim Valor
    Valor = Space(5)
    While Not AdoDbf.Recordset.EOF
        If AdoDbf.Recordset!pedirc > AdoDbf.Recordset!VENTAC Or AdoDbf.Recordset!pedirp > AdoDbf.Recordset!ventaP Then
           RSet Valor = AdoDbf.Recordset!pedirc
           Print #1, Mid(AdoDbf.Recordset!descripc, 1, 35) & "   " & AdoDbf.Recordset!medida & Valor;
           RSet Valor = AdoDbf.Recordset!VENTAC
           Print #1, Valor;
           RSet Valor = AdoDbf.Recordset!pedirc - AdoDbf.Recordset!VENTAC
           Print #1, Valor;
           
           RSet Valor = AdoDbf.Recordset!pedirp
           Print #1, " " & Valor;
           RSet Valor = AdoDbf.Recordset!ventaP
           Print #1, Valor;
           RSet Valor = AdoDbf.Recordset!pedirp - AdoDbf.Recordset!ventaP
           Print #1, Valor
        End If
        AdoDbf.Recordset.MoveNext
    Wend
    Close #1   ' Cierra el archivo de reporte
    Handle = Shell("NOTEPAD " & App.Path & "\FAL" & Mid(cArch, 4) & ".TXT", 1)
    Stb1.Panels(1).Text = cmen
    Stb1.Refresh
    Exit Sub
Error:
    If lTrans Then cn.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub cmdModAce_Click()
If Optsitua(0).Value = True Then
   Status = "1"
ElseIf Optsitua(1).Value = True Then
   Status = "2"
ElseIf Optsitua(2).Value = True Then
   Status = "3"
End If
If MsgBox("REALMENTE DESEAS CAMBIAR LA SITUACION DE LA VENTA A " & UCase(Optsitua(Val(Status) - 1).Caption) & Chr(13) & "DEL FOLIO UNICO DE LA VENTA " & AdoVentas.Recordset!noventa, vbYesNo + vbQuestion) = vbYes Then
'   MsgBox "UPDATE VENTAS SET SITUACION = " & Status & ",CORTEY1 = " & Me.txtcambia.Text & "  WHERE NOVENTA = " & AdoVentas.Recordset!NOVENTA
   cn.Execute "UPDATE VENTAS SET SITUACION = " & Status & ",CORTEY1 = " & Me.txtcambia.Text & "  WHERE NOVENTA = " & AdoVentas.Recordset!noventa
   AdoVentas.Refresh
End If
End Sub


Private Sub cmdModReg_Click()
  fracambia.Visible = False
End Sub

Private Sub cmdMoverse_Click(Index As Integer)
Dim rstBus As ADODB.Recordset
'On Error GoTo Error:
If AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF Then
   For N = 0 To 5
       Cmdmoverse(N).Enabled = False
   Next
   Exit Sub
End If
Select Case Index
Case 0  ' Primer registro
    AdoVentas.Recordset.MoveFirst
Case 1  ' Anterior
     frmCliente.Show
Case 2  ' Siguiente
     Me.AdoVentas.Refresh
Case 3  ' Ultimo
    AdoVentas.Recordset.MoveLast
Case 4  ' Buscar clave de la venta por dia
    cCve = InputBox("Introduzca el folio unico de la venta a buscar", "Introducir clave")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    Antes = dbgrdVta.Bookmark
    AdoVentas.Recordset.MoveFirst
    AdoVentas.Recordset.Find "Noventa = " & cCve
    If AdoVentas.Recordset.EOF Then
       MsgBox "LA CLAVE " & UCase(cCve) & " NO SE ENCUENTRA EN LAS VENTAS " + IIf(Me.cmbFiltro.Text = "TODAS", "" & Chr(13) & " EN EL PERIODO SELECCIONADO", cmbFiltro.Text), vbExclamation
       dbgrdVta.Bookmark = Antes
    End If
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdopcion_Click(Index As Integer)
Dim rs As ADODB.Recordset
Preventa = 0
Select Case Index
  Case 0 'Ventas
       nOp = 0
       frmVentas.Show
       frmVentas.Caption = "VENTAS A MAYOREO" & Space(10) & "COMPUTADORA: [" & Caja & "]" & Space(10) & "MOVIMIENTO: [Nueva venta]"
  Case 1 'Cobro de la venta
       nOp = 3
       frmVentas.Caption = "VENTAS A MAYOREO" & Space(10) & Space(10) & "COMPUTADORA: [" & Caja & "]" & Space(10) & "MOVIMIENTO: [Cobrar venta]"
       frmVentas.Show
  Case 2
       frmHistCred.Show 1
  Case 3
       'CANCELACION DE UNA VENTA PENDIENTE DE COBRAR, LA CANCELACION DE UNA FACTURA ES APARTE
       If (Me.AdoVentas.Recordset!situacion < 2) And (Me.AdoVentas.Recordset!cancelado = False) Then
            RESP = MsgBox("Realmente Desea Cancelar la venta...", vbYesNo, "CANCELACION DE VENTA")
            If RESP = vbYes Then
               Call cancelaventa
            End If
       Else
          MsgBox "No es posible cancelar una venta que ya fue cobrada , cancelada", vbInformation
       End If
  Case 6
        
  Case 4 ' Salir
       Unload Me
  Case 5 'Modificar venta
       nOp = 1
       frmVentas.Show
       frmVentas.Caption = "VENTAS A MAYOREO" & Space(10) & Space(10) & "COMPUTADORA: [" & Caja & "]" & Space(10) & "MOVIMIENTO: [Modificar venta]"
End Select

End Sub

Private Sub cancelaventa()
'el proceso regresa al inventario lo de cada venta
Me.Stb1.SimpleText = "Actualizando Inventario"
Me.AdoVentas.Recordset!cancelado = 1
AdoVentas.Recordset.Update
Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT * FROM ventas_det WHERE noventa  = " & AdoVentas.Recordset!noventa, cn, adOpenKeyset, adLockOptimistic, adCmdText
While Not RsCon.EOF
   CAD = "update inventario set incant = incant + " & RsCon!cantidad & " ,incantpza = incantpza +  " & RsCon!cantidadp & " where inprod = '" & Trim(RsCon!cl_producto) & "'"
   'MsgBox cad
   cn.Execute CAD
   RsCon.MoveNext
Wend
Me.AdoVentas.Recordset!situacion = 2
AdoVentas.Recordset!cancelado = 1
AdoVentas.Recordset.Update
Set RsCon = Nothing
End Sub


Private Sub cmdPend_Click()
frapend.Visible = True
AdoPend.ConnectionString = cCadConex
AdoPend.CommandType = adCmdText
AdoPend.RecordSource = "SELECT agente,d.noventa,v.fecha,folpreventa,descripc,LTRIM(STR(T.paquetes)) + ' x ' + LTRIM(STR(CONTENID,10,3)) + space(1) + RTRIM(t.medida) AS present,d.cantidad,d.precio,d.importe FROM Ventas v, ventas_det d, tfproduc t WHERE v.noventa =  d.noventa And consec = cl_producto And D.facturado = 0 " & cFecha & " Order by d.noventa"
AdoPend.Refresh
End Sub

Private Sub cmdtmp_Click()
Dim RSMDB As ADODB.Recordset
Dim CNMDB As ADODB.Connection
Dim Valor As String

Set RSMDB = New ADODB.Recordset
Set CNMDB = New ADODB.Connection
CNMDB.Open "DSN=PITICOMDB;DBQ=P:\PITICO\PITICO10\PITICO10.mdb;DefaultDir=P:\PITICO\PITICO10;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;PWD=PORTATIL"

AdoFranq.ConnectionTimeout = 0
AdoFranq.CommandTimeout = 0
AdoFranq.ConnectionString = cCadConex
RESP = MsgBox("Marcar productos como revisados en donde son igual la descripcion y presentacion en Oficinas y Miguel cabrera", vbYesNo + vbQuestion)
If RESP = vbYes Then
   'AdoFranq.RecordSource = "SELECT paquetes,CONSEC,descripc, LTRIM(STR(T.paquetes)) + ' x ' + LTRIM(STR(CONTENID,10,3)) + space(1) + RTRIM(t.medida) AS MEDIDA FROM inventario, tfproduc T WHERE T.consec = inprod AND incant > 0 and igualofi = 0 ORDER BY descripc, medida"
   AdoFranq.RecordSource = "SELECT paquetes,CONSEC,descripc, LTRIM(STR(T.paquetes)) + ' x ' + LTRIM(STR(CONTENID,10,3)) + space(1) + RTRIM(t.medida) AS MEDIDA FROM inventario, tfproduc T WHERE T.consec = inprod AND ACTIVO = 1 and igualofi = 0 ORDER BY descripc, medida"
Else
   AdoFranq.RecordSource = "SELECT paquetes,CONSEC,descripc, LTRIM(STR(T.paquetes)) + ' x ' + LTRIM(STR(CONTENID,10,3)) + space(1) + RTRIM(t.medida) AS MEDIDA FROM inventario, tfproduc T WHERE T.consec = inprod and igualofi = 1 ORDER BY descripc, medida"
End If
AdoFranq.Refresh
PRECIO = "Precio4"
While Not AdoFranq.Recordset.EOF
    Stb1.Panels(1).Text = AdoFranq.Recordset!CONSEC
    Stb1.Refresh
    RSMDB.Open "SELECT paquetes,precio4, precio1, descripc,LTRIM(STR(paquetes)) + ' x ' + LTRIM(STR(CONTENID)) + space(1) + RTRIM(medida) AS MEDIDA1  FROM tfproduc,preprod WHERE consec = preclave and consec = '" & AdoFranq.Recordset!CONSEC & "'", CNMDB, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not (RSMDB.BOF And RSMDB.EOF) Then
       If RESP = vbYes Then
          If AdoFranq.Recordset!PAQUETES = RSMDB!PAQUETES Then
             If MsgBox("MCABRERA.: " & AdoFranq.Recordset!descripc & " " & AdoFranq.Recordset!medida & Chr(13) & "OFICINAS:     " & RSMDB!descripc & " " & RSMDB!medida1, vbYesNo, "Comparativo Mcab y Oficinas") = vbYes Then
                cn.Execute "UPDATE tfproduc SET igualofi = 1 WHERE consec = '" & Me.AdoFranq.Recordset!CONSEC & "'"
             End If
          End If
       ElseIf AdoFranq.Recordset!PAQUETES <> RSMDB!PAQUETES Then
          If MsgBox("MCABRERA.: " & AdoFranq.Recordset!descripc & " " & AdoFranq.Recordset!medida & Chr(13) & "OFICINAS:     " & RSMDB!descripc & " " & RSMDB!medida1, vbYesNo + vbDefaultButton2, "Comparativo Mcab y Oficinas") = vbNo Then
             cn.Execute "UPDATE tfproduc SET igualofi = 0 WHERE consec = '" & Me.AdoFranq.Recordset!CONSEC & "'"
          End If
       End If
    End If
    RSMDB.Close
    AdoFranq.Recordset.MoveNext
Wend
MsgBox "NO EXISTEN MAS COINCIDENCIAS", vbInformation
End Sub

Private Function GRABCARBO(PRECAJ As Currency, PREPZA As Currency) As Boolean
On Error GoTo Error:
'AQUI SI SE DEBE IMPLEMENTAR UN RECORDSET PARA OBTENER EL FOLIO DE VENTA
AdoFranq.Recordset!PRECIO = PRECAJ
AdoFranq.Recordset!importe = (AdoFranq.Recordset!cantidad * PRECAJ) + (AdoFranq.Recordset!cantidadp * PREPZA)
AdoFranq.Recordset.Update
   GRABASERIEYFAC = True
Exit Function
Error:
   GRABASERIEYFAC = False
End Function

Private Sub Command1_Click()
  frapend.Visible = False
End Sub

Private Sub DBGRDPend_HeadClick(ByVal ColIndex As Integer)
campo = IIf(UCase(DBGRDPEND.Columns(ColIndex).DataField) = "NOVENTA", "D.NOVENTA", DBGRDPEND.Columns(ColIndex).DataField)
AdoPend.RecordSource = "SELECT agente,d.noventa,v.fecha,folpreventa,descripc,LTRIM(STR(T.paquetes)) + ' x ' + LTRIM(STR(CONTENID,10,3)) + space(1) + RTRIM(t.medida) AS present,d.cantidad,d.precio,d.importe FROM Ventas v, ventas_det d, tfproduc t WHERE v.noventa =  d.noventa And consec = cl_producto And D.facturado = 0 " & cFecha & " Order by " & campo & ", fecha"
AdoPend.Refresh
End Sub

Private Sub dbgrdVta_HeadClick(ByVal ColIndex As Integer)
  Stb1.SimpleText = "Espere un momento ordenando Pedidos por " & dbgrdVta.Columns(ColIndex).Caption
  AdoVentas.RecordSource = "SELECT * FROM Ventas,CatCliente WHERE cClave = clcliente AND " & cCond & cFecha & "ORDER BY " & dbgrdVta.Columns(ColIndex).DataField
  AdoVentas.Refresh
  Stb1.Panels(1).Text = Space(85) + "Ventas ordenandas por " & dbgrdVta.Columns(ColIndex).Caption
End Sub

Private Sub dbgrdVta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdopcion_Click 5
End Sub

Private Sub dtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtpFecha_LostFocus(Index As Integer)
On Error GoTo Error:
 
 If Sql Then
    cFecha = " AND fecha >= '" & Format(dtpFecha(0).Value, "DD-MM-YYYY") & "' AND fecha <= '" & Format(DateAdd("d", 1, dtpFecha(1).Value), "DD-MM-YYYY") & "'"         'Cargo todas las ventas
 Else
    cFecha = " AND fecha >= #" & Format(dtpFecha(0).Value, "MM-DD-YYYY") & "# AND fecha <= #" & Format(DateAdd("d", 1, dtpFecha(1).Value), "MM-DD-YYYY") & "#"         'Cargo todas las ventas
 End If
 
 AdoVentas.RecordSource = "SELECT * FROM Ventas, CatCliente WHERE clcliente = cClave AND " & cCond & cFecha & " ORDER BY cNombre"
 AdoVentas.Refresh
 lblInfo.Caption = Str(AdoVentas.Recordset.RecordCount)
 For N = 0 To 4   'Si esta vacio el recordset desactivo las opciones
   Cmdmoverse(N).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
 Next
 cmdopcion(0).Enabled = ModVta
 cmdopcion(1).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
 cmdopcion(5).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
 Exit Sub
Error:
  MsgBox Err.Description

End Sub

Private Sub fechacambio_LostFocus()
'SE ACTUALIZA LA FECHA DE LA VENTA
AdoVentas.Recordset!fecha = fechacambio.Value
AdoVentas.Recordset.Update
MsgBox "El cambio de Fecha se ha realizado Correctamente", vbInformation, "FECHA"
fechacambio.Enabled = False
fechacambio.Visible = False
Me.dbgrdVta.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
  frmPrincipal.Caption = "VENTAS A MAYOREO" & Space(10) & "COMPUTADORA: [" & CStr(Caja) & "]" & Space(15) & "USUARIO:[" + Mid(cCveDesUsu, 3) + "]" + Space(15) + "SUCURSAL:[" + Mid(cSucursal, 3) + "]"
  If Preventa = False Then txtcliente.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 115
       nOp = 3
       frmVentas.Caption = "VENTAS A MAYOREO" & Space(10) & Space(10) & "COMPUTADORA: [" & Caja & "]" & Space(10) & "MOVIMIENTO: [Cobrar venta]"
       frmVentas.Show
  Case 119
     frmCalc.Show 1
  Case 120
        nConPrin = 0
        Me.fraCon.Visible = True
        Me.txtContra.Text = ""
        Me.txtContra.SetFocus
  Case 121
       nConPrin = 1
       Me.fraCon.Visible = True
       Me.txtContra.Text = ""
       Me.txtContra.SetFocus
  Case 122     'Tecla de función [F11]
       If Trim(Mid(cSucursal, 1, 2)) = "55" Then
       If (AdoVentas.Recordset!clcliente = 11529 Or AdoVentas.Recordset!clcliente = 12953 Or AdoVentas.Recordset!clcliente = 12007 Or AdoVentas.Recordset!clcliente = 12145) Then
          If IsNull(AdoVentas.Recordset!tiket) Then
             cn.Execute "UPDATE ventas_det SET preciop = ROUND( preciop  - (preciop * 0.02) ,1,-2) + .10, precio = ROUND( precio  - (precio * 0.02) ,1,-2) + .10 FROM tfproduc t WHERE t.consec = cl_producto And NOT t.descripc like '%CIGARRO%' AND NOT T.claprove IN ('C66','C98','P17','P12','L34','C73') and noventa = " & AdoVentas.Recordset!noventa
             'cn.Execute "UPDATE ventas_det SET preciop = ROUND( preciop  - (preciop * 0.01) ,1,-2) + .10, precio = ROUND( precio - (precio * 0.01) ,1,-2) + .10 FROM tfproduc t WHERE t.consec = cl_producto And t.descripc like '%CIGARRO%' and noventa = " & AdoVentas.Recordset!noventa
             cn.Execute "UPDATE ventas_det SET importe = (PRECIO * cantidad) + (preciop * cantidadp) WHERE noventa = " & AdoVentas.Recordset!noventa
             cn.Execute "UPDATE ventas SET tiket = 'AJU' WHERE noventa = " & AdoVentas.Recordset!noventa
             MsgBox "LA ACTUALIZACION SE REALIZO CORRECTAMENTE", vbInformation, "Ventas"
          Else
             MsgBox "A ESTE CLIENTE YA SE LE ACTUALIZARON PRECIOS", vbExclamation, "Ventas"
          End If
       Else
          MsgBox "Este cliente no esta autorizado para precio especial", vbInformation, "Sin autorización"
       End If
       End If
  Case 123
       
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Form_Load()
  Unload frmLogin
  cmbFiltro.AddItem "TODAS"
  cmbFiltro.AddItem "EN TRAMITE CONTADO"
  cmbFiltro.AddItem "EN TRAMITE CREDITO-PREVTA"
  cmbFiltro.AddItem "COBRADAS"
  cmbFiltro.ListIndex = 0
   
  Set rstCli = New ADODB.Recordset
  rstCli.Open "SELECT * FROM CATCLIENTE", cn, adOpenKeyset, adLockOptimistic, adCmdText
  cCond = "NOVENTA > 0 "
  If dtpFecha(0).Value = "01/01/01" Then dtpFecha(0).Value = Format(date, "DD/mm/yyyy")
  If dtpFecha(1).Value = "01/01/01" Then dtpFecha(1).Value = Format(date, "DD/mm/yyyy")
  
  If Sql Then
     cFecha = " AND fecha >= '" & Format(dtpFecha(0).Value, "DD-MM-YYYY") & "' AND (fecha) <= '" & Format(DateAdd("d", 1, dtpFecha(1).Value), "DD-MM-YYYY") & "'"
  Else
     cFecha = " AND fecha >= #" & Format(dtpFecha(0).Value, "MM-DD-YYYY") & "# AND (fecha) <= #" & Format(DateAdd("d", 1, dtpFecha(1).Value), "MM-DD-YYYY") & "#"
  End If
  AdoVentas.ConnectionString = cCadConex
  AdoVentas.CommandType = adCmdText
  AdoVentas.RecordSource = "SELECT * FROM Ventas, CatCliente WHERE clcliente = cClave AND " & cCond & cFecha & " ORDER BY cNombre"
  AdoVentas.Refresh
  For N = 0 To 4
     Cmdmoverse(N).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
  Next
  
  'cmdopcion(0).Enabled = ModVta
  cmdopcion(1).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
  cmdopcion(5).Enabled = Not (AdoVentas.Recordset.BOF And AdoVentas.Recordset.EOF)
  lblInfo.Caption = Str(AdoVentas.Recordset.RecordCount)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim controlADO As Adodc
'Dim controlADO As Object
'Set controlADO = New ADODB.Recordset
' For Each controlADO In Adodc.Controls
'     Adodc.Name
' Next
 
' frmMayoreo.Show
 frmAreaRecibo.Show
End Sub

Private Sub mnucancela_Click()
cmdopcion_Click 3
End Sub

Private Sub mnucobro_Click()
cmdopcion_Click 1
End Sub

Private Sub mnufacvta_Click()
nOp = 1
frmVentas.Show
frmVentas.Caption = "VENTAS A MAYOREO" & Space(10) & Space(10) & "COMPUTADORA: [" & Caja & "]" & Space(10) & "MOVIMIENTO: [Modificar venta]"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
Exit Sub
lpprov = False
fconfac.Show 1
If lpprov Then
          'SE DETERMINAN CUANTOS PRODUCTOS SON EN LA VENTA Y CON BASE A ESO SE GENERAN N FACTURAS
          totfacturas = 2
          Call agregarfactura(fconfac.txtafactura.Text, fconfac.txtaserie.Text, fconfac.cmbCliente.Text)
          MsgBox "Proceso Finalizado...", vbInformation
Else
          MsgBox "No es posible Generar La factura si ya existe", vbInformation
End If
End Sub

Private Sub mnufecha_Click()
fechacambio.Enabled = True
fechacambio.Visible = True
fechacambio.SetFocus
End Sub

Private Sub mnuhisto_Click()
cmdopcion_Click 2
End Sub

Private Sub mnumodi_Click()
cmdopcion_Click 5
End Sub

Private Sub mnunueva_Click()
cmdopcion_Click 0
End Sub

Private Sub mnusal_Click()
Unload Me
End Sub

Private Sub txtcliente_Change()
On Error Resume Next
  'MsgBox "cnombre = '*" & Me.txtcliente.Text & "'"
  AdoVentas.Recordset.MoveFirst
  AdoVentas.Recordset.Find "cnombre LIKE '" & Me.txtcliente.Text & "*'"
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmdborravta_Click()
Dim RST As ADODB.Recordset
Set RST = New ADODB.Recordset
If MsgBox("Confirma si deseas borrar la venta con folio unico " & AdoPend.Recordset!noventa, vbYesNo + vbQuestion, "Ventas") = vbYes Then
   RST.Open "SELECT * FROM VENTAS_DET WHERE noventa = " & AdoPend.Recordset!noventa, cn, adOpenDynamic, adLockOptimistic, adCmdText
   men = Stb1.Panels(1).Text
   While Not RST.EOF
      Stb1.Panels(1).Text = "Procesando: " & RST!cl_producto
      Stb1.Refresh
      cn.Execute "DELETE FROM Ventas_det WHERE cl_producto = '" & RST!cl_producto & "' AND noventa = " & RST!noventa
      cn.Execute "UPDATE Inventario SET InCant = Incant + " & RST!cantidad & ", InCantPza = IncantPza + " & RST!cantidadp & " WHERE Inprod = '" & RST!cl_producto & "'"
      RST.MoveNext
   Wend
   Stb1.Panels(1).Text = men
   RST.Close
   Set RST = Nothing
   cn.Execute "UPDATE VEnTAS SET montototal = 0 where noventa = " & AdoPend.Recordset!noventa
   marca = AdoPend.Recordset.Bookmark
   AdoPend.Refresh
   AdoPend.Recordset.Bookmark = marca
   MsgBox "LA VENTA SE BORRO CORRECTAMENTE", vbInformation, "Ventas"
End If
End Sub
