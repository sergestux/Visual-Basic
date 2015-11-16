VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frminvtodo 
   Caption         =   "Inventario Global de Bodegas de Mayoreo de Viveres y Licores S.A. de C.V."
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   Icon            =   "frminvtodo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
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
      Height          =   7695
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   11625
      Begin VB.TextBox txtpv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   9120
         TabIndex        =   54
         Top             =   4800
         Width           =   1800
      End
      Begin VB.TextBox txtcosto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   7200
         TabIndex        =   53
         Top             =   4800
         Width           =   1800
      End
      Begin VB.TextBox ttxvari 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   5400
         TabIndex        =   52
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtpzas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   4080
         TabIndex        =   51
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtcajas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   2640
         TabIndex        =   50
         Top             =   4800
         Width           =   1215
      End
      Begin VB.PictureBox CR1 
         Height          =   480
         Left            =   840
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   56
         Top             =   1320
         Width           =   1200
      End
      Begin VB.TextBox txtcajas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   2640
         TabIndex        =   47
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtpzas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   4080
         TabIndex        =   46
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox ttxvari 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   5400
         TabIndex        =   45
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtcosto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   7200
         TabIndex        =   44
         Top             =   4200
         Width           =   1800
      End
      Begin VB.TextBox txtpv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   9120
         TabIndex        =   43
         Top             =   4200
         Width           =   1800
      End
      Begin VB.TextBox txtpv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   9120
         TabIndex        =   33
         Top             =   3600
         Width           =   1800
      End
      Begin VB.TextBox txtpv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   9120
         TabIndex        =   32
         Top             =   3000
         Width           =   1800
      End
      Begin VB.TextBox txtpv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   9120
         TabIndex        =   31
         Top             =   2400
         Width           =   1800
      End
      Begin VB.TextBox txtcosto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   7200
         TabIndex        =   30
         Top             =   3600
         Width           =   1800
      End
      Begin VB.TextBox txtcosto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   7200
         TabIndex        =   29
         Top             =   3000
         Width           =   1800
      End
      Begin VB.TextBox txtcosto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   7200
         TabIndex        =   28
         Top             =   2400
         Width           =   1800
      End
      Begin VB.TextBox ttxvari 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   5400
         TabIndex        =   27
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox ttxvari 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   5400
         TabIndex        =   26
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox ttxvari 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   5400
         TabIndex        =   25
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtpzas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   4080
         TabIndex        =   24
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtpzas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   4080
         TabIndex        =   23
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtpzas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   4080
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtcajas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   2640
         TabIndex        =   21
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtcajas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   2640
         TabIndex        =   20
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtcajas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   2640
         TabIndex        =   19
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Volver"
         Height          =   495
         Left            =   9000
         Picture         =   "frminvtodo.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6600
         Width           =   1575
      End
      Begin VB.TextBox ttxvari 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   5400
         TabIndex        =   17
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtpv 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   9120
         TabIndex        =   16
         Top             =   5760
         Width           =   1800
      End
      Begin VB.TextBox txtcosto 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   7200
         TabIndex        =   15
         Top             =   5760
         Width           =   1800
      End
      Begin VB.TextBox txtpzas 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   4080
         TabIndex        =   14
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtcajas 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Istmo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   240
         TabIndex        =   55
         Top             =   4800
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Miahuatlan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   240
         TabIndex        =   48
         Top             =   4200
         Width           =   1995
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio de venta"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   9120
         TabIndex        =   42
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio de costo"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   7200
         TabIndex        =   41
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Variedades"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   5400
         TabIndex        =   40
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Piezas"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   4200
         TabIndex        =   39
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajas"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   38
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblrtiquetas 
         Alignment       =   2  'Center
         Caption         =   "INVENTARIOS DE LAS BODEGAS DE MAYOREO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   10815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Puerto Escondido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Central de Abastos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   3000
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Miguel Cabrera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Mayoreo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   5760
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frminvtodo.frx":05B4
         Height          =   7335
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12632256
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   17
         RowDividerStyle =   4
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
         Caption         =   "EXISTECIA DE PRODUCTOS DE LAS BODEGAS DE MAYOREO"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adoinv 
      Height          =   330
      Left            =   240
      Top             =   360
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
   Begin VB.PictureBox PicBotones 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   570
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   10320
      TabIndex        =   1
      Top             =   6285
      Width           =   10380
      Begin VB.CommandButton cmdrpt 
         Caption         =   "&Reporte"
         Height          =   450
         Left            =   3480
         Picture         =   "frminvtodo.frx":05C9
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Regresar al menu principal"
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Totales"
         Height          =   450
         Left            =   2640
         Picture         =   "frminvtodo.frx":0AFB
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "totales"
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   450
         Left            =   4320
         Picture         =   "frminvtodo.frx":0BFD
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Regresar al menu principal"
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton cmdBuscaDesc 
         Caption         =   "&Prod"
         Height          =   450
         Left            =   960
         Picture         =   "frminvtodo.frx":0D6F
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Busqueda por descripcion"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton CmdExporta 
         Caption         =   "&Excel"
         Height          =   450
         Left            =   1800
         Picture         =   "frminvtodo.frx":0E69
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exportar existencias a archivo DBF"
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton cmdBuscaBarra 
         Caption         =   "&Clave"
         Height          =   450
         Left            =   120
         Picture         =   "frminvtodo.frx":0F6B
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Busqueda por codigo de barras"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chktienda 
         Caption         =   "Tienda"
         Height          =   220
         Left            =   7080
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chktodos 
         Caption         =   "Todos"
         Height          =   220
         Left            =   7080
         TabIndex        =   2
         Top             =   270
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   8520
         TabIndex        =   8
         Top             =   120
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frminvtodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscaBarra_Click()
Dim cCve As String
Dim Antes
cCve = InputBox("Introduzca el código del producto a buscar", "Introducir código")
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = DataGrid1.Bookmark
Adoinv.Recordset.MoveFirst
Adoinv.Recordset.Find "clave =" & Trim(cCve)
If Adoinv.Recordset.EOF Then
   MsgBox "El código " & cCve & " no se encuentra en el inventario", vbExclamation
   DataGrid1.Bookmark = Antes
End If
DataGrid1.SetFocus

End Sub

Private Sub cmdBuscaDesc_Click()
Dim cCve As String
Dim Antes
On Error GoTo Error:
cCve = InputBox("Introduzca la descripcion del producto a buscar", "Introducir descripcion")
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = DataGrid1.Bookmark
'me.
Adoinv.Recordset.MoveFirst
Adoinv.Recordset.Find "prod LIKE '" & Trim(cCve) & "*'"
If Adoinv.Recordset.EOF Then
   MsgBox "La descripcion " & cCve & " no se encuentra en el inventario", vbExclamation
   DataGrid1.Bookmark = Antes
End If
DataGrid1.SetFocus
Exit Sub
Error:
  Exit Sub
End Sub

Private Sub cmdRegresar_Click()
Unload Me
frmModInv.SetFocus
End Sub

Private Sub cmdRpt_Click()
cr1.ReportFileName = App.Path & "\invbode.rpt"
cr1.Connect = cCadConex
If frmModInv.txtclave.Text = "" Then
   cconprov = ""
   cProv = ""
Else
   cconprov = " AND tfproduc.claprove = '" & Trim(frmModInv.txtclave.Text) & "'"
   cProv = " DE " & frmModInv.cmbProved.Text
End If
cconprov = IIf(frmModInv.txtclave.Text = "", "", " AND tfproduc.claprove = '" & Trim(frmModInv.txtclave.Text) & "'")
cr1.SQLQuery = "SELECT INVENTARIO.inprod, INVENTARIO.incant, " & _
                       "INVENTARIO55.incant,TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.activo, " & _
                       "INVENTARIO23.incant, INVENTARIO26.incant,INVENTARIO28.incant " & Chr(13) & _
               "FROM pitico.dbo.INVENTARIO INVENTARIO,PITICO.dbo.INVENTARIO55 INVENTARIO55, " & _
                     "PITICO.dbo.TFPRODUC TFPRODUC, " & _
                     "PITICO.dbo.INVENTARIO23 INVENTARIO23, pitico.dbo.INVENTARIO26 INVENTARIO26, pitico.dbo.INVENTARIO28 INVENTARIO28 " & Chr(13) & _
               "WHERE INVENTARIO.inprod = INVENTARIO55.inprod AND " & _
                     "INVENTARIO.inprod = TFPRODUC.CONSEC AND " & _
                     "INVENTARIO55.inprod = INVENTARIO23.inprod AND " & _
                     "INVENTARIO23.inPROD = INVENTARIO26.inPROD AND INVENTARIO26.INPROD = INVENTARIO28.INPROD AND " & _
                     "TFPRODUC.activo = 1" & cconprov
cEnca = "INVENTARIOS DE LAS BODEGAS DE MAYOREO " & cProv
cr1.Formulas(0) = "ENCABEZADO = '" & cEnca & "'"
cr1.WindowTitle = cEnca
cr1.Action = 1
End Sub

Private Sub Command1_Click()
Frame2.Enabled = True
Frame2.Visible = True
Call suma
Frame2.Enabled = True
Frame2.Visible = True
End Sub


Private Sub suma()
For i = 0 To 5
   txtcajas(i).Text = 0
   txtpzas(i).Text = 0
   ttxvari(i).Text = 0
   txtcosto(i).Text = 0
   txtpv(i).Text = 0
Next
Adoinv.Recordset.MoveFirst
While Not Adoinv.Recordset.EOF
    VPV = VPV + 1
    'TOTAL DE CAJAS
    txtcajas(0).Text = txtcajas(0).Text + IIf(Not IsNull(Adoinv.Recordset!totcaj), Adoinv.Recordset!totcaj, 0)
    txtcajas(1).Text = txtcajas(1).Text + IIf(Not IsNull(Adoinv.Recordset!MC_caj), Adoinv.Recordset!MC_caj, 0)
    txtcajas(2).Text = txtcajas(2).Text + IIf(Not IsNull(Adoinv.Recordset!CA_caj), Adoinv.Recordset!CA_caj, 0)
    txtcajas(3).Text = txtcajas(3).Text + IIf(Not IsNull(Adoinv.Recordset!PE_caj), Adoinv.Recordset!PE_caj, 0)
    txtcajas(4).Text = txtcajas(4).Text + IIf(Not IsNull(Adoinv.Recordset!MH_caj), Adoinv.Recordset!MH_caj, 0)
    txtcajas(5).Text = txtcajas(5).Text + IIf(Not IsNull(Adoinv.Recordset!IS_caj), Adoinv.Recordset!IS_caj, 0)
    'TOTAL DE PIEZAS
    txtpzas(0).Text = txtpzas(0).Text + IIf(Not IsNull(Adoinv.Recordset!totpza), Adoinv.Recordset!totpza, 0)
    txtpzas(1).Text = txtpzas(1).Text + IIf(Not IsNull(Adoinv.Recordset!MC_pza), Adoinv.Recordset!MC_pza, 0)
    txtpzas(2).Text = txtpzas(2).Text + IIf(Not IsNull(Adoinv.Recordset!CA_pza), Adoinv.Recordset!CA_pza, 0)
    txtpzas(3).Text = txtpzas(3).Text + IIf(Not IsNull(Adoinv.Recordset!PE_PZA), Adoinv.Recordset!PE_PZA, 0)
    txtpzas(4).Text = txtpzas(4).Text + IIf(Not IsNull(Adoinv.Recordset!MH_PZA), Adoinv.Recordset!MH_PZA, 0)
    txtpzas(5).Text = txtpzas(5).Text + IIf(Not IsNull(Adoinv.Recordset!IS_PZA), Adoinv.Recordset!IS_PZA, 0)
    'On Error Resume Next
    'COSTO POR TIENDA
    txtcosto(0).Text = txtcosto(0).Text + (Adoinv.Recordset!PRECOSTO * IIf(IsNull(Adoinv.Recordset!totcaj), 0, Adoinv.Recordset!totcaj)) + (Adoinv.Recordset!PRECOSTO / Adoinv.Recordset!paq * IIf(IsNull(Adoinv.Recordset!totpza), 0, Adoinv.Recordset!totpza))
    txtcosto(1).Text = txtcosto(1).Text + (Adoinv.Recordset!PRECOSTO * Adoinv.Recordset!MC_caj) + (Adoinv.Recordset!PRECOSTO / Adoinv.Recordset!paq * Adoinv.Recordset!MC_pza)
    txtcosto(2).Text = txtcosto(2).Text + (Adoinv.Recordset!PRECOSTO * Adoinv.Recordset!CA_caj) + (Adoinv.Recordset!PRECOSTO / Adoinv.Recordset!paq * Adoinv.Recordset!CA_pza)
    txtcosto(3).Text = txtcosto(3).Text + (Adoinv.Recordset!PRECOSTO * IIf(IsNull(Adoinv.Recordset!PE_caj), 0, Adoinv.Recordset!PE_caj)) + (Adoinv.Recordset!PRECOSTO / Adoinv.Recordset!paq * IIf(IsNull(Adoinv.Recordset!PE_PZA), 0, Adoinv.Recordset!PE_PZA))
    txtcosto(4).Text = txtcosto(4).Text + (Adoinv.Recordset!PRECOSTO * IIf(IsNull(Adoinv.Recordset!MH_caj), 0, Adoinv.Recordset!MH_caj)) + (Adoinv.Recordset!PRECOSTO / Adoinv.Recordset!paq * IIf(IsNull(Adoinv.Recordset!MH_PZA), 0, Adoinv.Recordset!MH_PZA))
    txtcosto(5).Text = txtcosto(5).Text + (Adoinv.Recordset!PRECOSTO * IIf(IsNull(Adoinv.Recordset!IS_caj), 0, Adoinv.Recordset!IS_caj)) + (Adoinv.Recordset!PRECOSTO / Adoinv.Recordset!paq * IIf(IsNull(Adoinv.Recordset!IS_PZA), 0, Adoinv.Recordset!IS_PZA))
    
    'VENTA POR TIENDA
    txtpv(0).Text = txtpv(0).Text + (Adoinv.Recordset!precio4 * Adoinv.Recordset!totcaj) + (Adoinv.Recordset!precio1 * Adoinv.Recordset!totpza)
    txtpv(1).Text = txtpv(1).Text + (Adoinv.Recordset!precio4 * Adoinv.Recordset!MC_caj) + (Adoinv.Recordset!precio1 * Adoinv.Recordset!MC_pza)
    txtpv(2).Text = txtpv(2).Text + (Adoinv.Recordset!precio4 * Adoinv.Recordset!CA_caj) + (Adoinv.Recordset!precio1 * Adoinv.Recordset!CA_pza)
    txtpv(3).Text = txtpv(3).Text + (Adoinv.Recordset!precio4 * Adoinv.Recordset!PE_caj) + (Adoinv.Recordset!precio1 * Adoinv.Recordset!PE_PZA)
    txtpv(4).Text = txtpv(4).Text + (Adoinv.Recordset!precio4 * Adoinv.Recordset!MH_caj) + (Adoinv.Recordset!precio1 * Adoinv.Recordset!MH_PZA)
    txtpv(5).Text = txtpv(5).Text + (Adoinv.Recordset!precio4 * Adoinv.Recordset!IS_caj) + (Adoinv.Recordset!precio1 * Adoinv.Recordset!IS_PZA)
    'VARIEDADES POR TIENDA
    If Adoinv.Recordset!totcaj > 0 Then ttxvari(0).Text = Me.ttxvari(0).Text + 1
    If Adoinv.Recordset!MC_caj > 0 Then ttxvari(1).Text = Me.ttxvari(1).Text + 1
    If Adoinv.Recordset!CA_caj > 0 Then ttxvari(2).Text = Me.ttxvari(2).Text + 1
    If Adoinv.Recordset!PE_caj > 0 Then ttxvari(3).Text = Me.ttxvari(3).Text + 1
    If Adoinv.Recordset!MH_caj > 0 Then ttxvari(4).Text = Me.ttxvari(4).Text + 1
    If Adoinv.Recordset!IS_caj > 0 Then ttxvari(5).Text = Me.ttxvari(5).Text + 1
    Adoinv.Recordset.MoveNext
Wend
cajas = 0
pzas = 0
costo = 0
pv1 = 0
For i = 1 To 5
  cajas = cajas + Val(txtcajas(i).Text)
  pzas = pzas + Val(txtpzas(i).Text)
  costo = costo + Val(txtcosto(i).Text)
  pv1 = pv1 + Val(txtpv(i).Text)
Next
txtcajas(0).Text = cajas
txtpzas(0).Text = pzas
txtcosto(0).Text = costo
txtpv(0).Text = pv1
'SE COMPONEN LOS NUMEROS EN SU FORMATO
For i = 0 To 5
  txtpv(i).Text = Format(Round(Val(txtpv(i).Text), 2), "$ ###,###,###.00")
  txtcosto(i).Text = Format(Round(Val(txtcosto(i).Text), 2), "$ ###,###,###.00")
Next
Me.Frame2.Refresh

End Sub
Private Sub Command2_Click()
Me.Frame2.Enabled = False
Frame2.Visible = False
End Sub

Private Sub Form_Load()
'On Error GoTo Error:
 Adoinv.CursorType = adOpenKeyset
 Adoinv.ConnectionString = cCadConex
 'tiendas por orden de prioridad
 cconprov = IIf(frmModInv.txtclave.Text = "", "", " AND claprove = '" & Trim(frmModInv.txtclave.Text) & "'")
 CAD = "SELECT consec as clave ,descripc prod, LTRIM(RTRIM(STR(PAQUETES))) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' + MEDIDA AS PRESENT , " & _
       " (i16.incant + i23.incant + i55.incant + I26.incant + I28.incant) AS TotCaj,  " & _
       " (i16.incantpza + i23.incantpza + i55.incantpza + I26.incantpza + I28.incantpza) AS Totpza,  " & _
       " I16.INCANT AS MC_Caj,I16.INCANTPZA AS MC_Pza," & _
       " I23.INCANT AS CA_Caj,I23.INCANTPZA AS CA_Pza," & _
       " I55.INCANT AS PE_Caj,I55.INCANTPZA AS PE_Pza," & _
       " I26.INCANT AS MH_Caj,I26.INCANTPZA AS MH_Pza," & _
       " I28.INCANT AS IS_Caj,I28.INCANTPZA AS IS_Pza, PRECOSTO,PRECIO1,PRECIO4, paquetes paq " & _
       " FROM tfproduc,preprod,inventario as i16, inventario23 as i23, inventario55 as i55, Inventario26 as I26, inventario28 as I28  " & _
       " WHERE CONSEC = PRECLAVE and consec = i16.inprod And consec = i23.inprod AND consec = i55.inprod AND consec = I26.inprod AND consec = I28.inprod AND (activo = 1 or (i16.incant + i23.incant + i55.incant + I26.incant + I28.incant) > 0 ) " & cconprov & _
       " ORDER BY prod "
'MsgBox CAD
Adoinv.RecordSource = CAD
Adoinv.Refresh
lblInfo.Caption = "Variedades de Productos:  " & Adoinv.Recordset.RecordCount
DataGrid1.Columns(0).Width = 800
DataGrid1.Columns(1).Width = 4500
DataGrid1.Columns(2).Width = 1500
DataGrid1.Columns(3).Width = 850
DataGrid1.Columns(4).Width = 0
For i = 5 To DataGrid1.Columns.Count - 1
   DataGrid1.Columns(i).Width = 800
   DataGrid1.Columns(i).Visible = (i Mod 2 = 1)
   DataGrid1.Columns(i).Alignment = dbgRight
Next
DataGrid1.Refresh
Me.Adoinv.Recordset.Find "CLAVE = '" & frmModInv.AdoModInv.Recordset!Inprod & "'"
Exit Sub
Error:
   MsgBox "No se ha podido completar las existencias Globales, verifique que la base tenga a todas las tiendas", vbInformation, "INVENTARIOS"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAreaRecibo.Show
End Sub


