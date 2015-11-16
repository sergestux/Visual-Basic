VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fprov 
   Caption         =   "Catálogo de Proveedores"
   ClientHeight    =   8595
   ClientLeft      =   2190
   ClientTop       =   1515
   ClientWidth     =   9660
   Icon            =   "frmprov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Proveedor"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   11655
      Begin VB.CheckBox chkvolumen 
         Caption         =   "Volúmen"
         DataField       =   "volumen"
         DataSource      =   "adoprov"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   9360
         TabIndex        =   90
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cmbcomprador 
         Height          =   315
         Left            =   1800
         TabIndex        =   68
         Top             =   3240
         Width           =   3615
      End
      Begin VB.CheckBox Chkback 
         Caption         =   "BackOrder"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   7320
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.ListBox LstCargo 
         Height          =   1035
         ItemData        =   "frmprov.frx":0442
         Left            =   9480
         List            =   "frmprov.frx":0455
         TabIndex        =   62
         Top             =   5280
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Txtvisita 
         Height          =   315
         Left            =   9840
         TabIndex        =   13
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox Chktipop 
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5640
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Txtrfc 
         Height          =   315
         Left            =   8400
         MaxLength       =   13
         TabIndex        =   5
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox Chkactivo 
         Caption         =   "Activo"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Txtprove 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   0
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Cmbtipop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmprov.frx":04B0
         Left            =   2040
         List            =   "frmprov.frx":04C0
         TabIndex        =   12
         Text            =   "Indirecto"
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox Txttelpro 
         Height          =   315
         Left            =   6720
         MaxLength       =   40
         TabIndex        =   11
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox Txtcodpro 
         Height          =   315
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   10
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Txtciupro 
         Height          =   315
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox Txtdelpro 
         Height          =   315
         Left            =   1800
         MaxLength       =   26
         TabIndex        =   8
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox txtcolpro 
         Height          =   315
         Left            =   7800
         MaxLength       =   25
         TabIndex        =   7
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox Txtfrecuencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9840
         MaxLength       =   5
         TabIndex        =   14
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtdirpro 
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtnomprove 
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label lblproina 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Productos activos: XX"
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
         Left            =   600
         TabIndex        =   89
         Top             =   360
         Width           =   3300
      End
      Begin VB.Label lblproAct 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Productos activos: XX"
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
         Left            =   4200
         TabIndex        =   88
         Top             =   360
         Width           =   3300
      End
      Begin VB.Label Label16 
         Height          =   375
         Left            =   8040
         TabIndex        =   64
         Top             =   4440
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "Comprador:"
         Height          =   255
         Left            =   600
         TabIndex        =   67
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Lblactivo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   7800
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.Label Label15 
         Caption         =   "Periodo de Visitas en Dias "
         Height          =   255
         Left            =   7680
         TabIndex        =   61
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "R.F.C."
         Height          =   255
         Left            =   7560
         TabIndex        =   59
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Clave:"
         Height          =   255
         Left            =   600
         TabIndex        =   55
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo de Proveedor"
         Height          =   255
         Left            =   600
         TabIndex        =   54
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Días que transcurren para entrega de pedido"
         Height          =   255
         Left            =   5880
         TabIndex        =   45
         Top             =   3840
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   5760
         TabIndex        =   44
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Código Postal"
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   6480
         TabIndex        =   42
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Delegación"
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Colonia"
         Height          =   255
         Left            =   6840
         TabIndex        =   40
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre "
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   1320
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adolinxpro 
      Height          =   330
      Left            =   5880
      Top             =   -120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc Adodcusu 
      Height          =   330
      Left            =   2880
      Top             =   -120
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
      Enabled         =   0
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
   Begin MSAdodcLib.Adodc adolinea 
      Height          =   375
      Left            =   1080
      Top             =   -120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "pitico"
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
   Begin MSAdodcLib.Adodc Adorepre 
      Height          =   330
      Left            =   3960
      Top             =   3120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   ""
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   7440
      Width           =   11655
      Begin VB.CommandButton cmdrep 
         Caption         =   "Reptante."
         Height          =   400
         Left            =   5040
         TabIndex        =   69
         ToolTipText     =   "Ver representante del proveedor"
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdlin 
         Caption         =   "&Asignar Familias"
         Height          =   400
         Left            =   3720
         TabIndex        =   47
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton Command6 
         Height          =   400
         Left            =   2640
         Picture         =   "frmprov.frx":04ED
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   400
         Left            =   8280
         TabIndex        =   26
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Command4 
         Height          =   400
         Left            =   2040
         Picture         =   "frmprov.frx":05E7
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Ir al ultimo registro"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton Command3 
         Height          =   400
         Left            =   1440
         Picture         =   "frmprov.frx":0759
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ir al siguiente registro"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton Command2 
         Height          =   400
         Left            =   840
         Picture         =   "frmprov.frx":08CB
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Ir al registro anterior"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton Command1 
         Height          =   400
         Left            =   240
         Picture         =   "frmprov.frx":0A3D
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ir al primer registro"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdcancela 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   400
         Left            =   9360
         TabIndex        =   33
         Top             =   240
         Width           =   1000
      End
      Begin MSAdodcLib.Adodc adoprov 
         Height          =   450
         Left            =   0
         Top             =   1680
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   794
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
         ConnectStringType=   1
         Appearance      =   0
         BackColor       =   -2147483645
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
         Caption         =   ""
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
      Begin VB.CommandButton cmdsalir 
         Caption         =   "Regresar"
         Height          =   400
         Left            =   10440
         Picture         =   "frmprov.frx":0BAF
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Salir del modulo"
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdmodifica 
         Caption         =   "&Modificar"
         Height          =   400
         Left            =   6120
         TabIndex        =   31
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   400
         Left            =   7200
         TabIndex        =   30
         Top             =   240
         Width           =   1000
      End
      Begin MSAdodcLib.Adodc Adousuario 
         Height          =   330
         Left            =   3000
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
   End
   Begin VB.Frame fradectos 
      Caption         =   "Descuentos financieros"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   77
      Top             =   4440
      Width           =   11655
      Begin VB.CommandButton cmddectofin 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Actualiza Descuentos Financieros de todos sus productos"
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   9975
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   7560
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   7560
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   7560
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   7560
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   7560
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Observaciones decto. Financiero 2"
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   87
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Observaciones decto. Financiero 3"
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   86
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Observaciones decto. Financiero 4"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   85
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Observaciones decto. Financiero 5"
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   84
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Observaciones decto. Financiero 1"
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   83
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Descuento Financiero 2"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   82
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Descuento Financiero 3"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   81
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Descuento Financiero 4"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   80
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Descuento Financiero 5"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   79
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Descuento Financiero 1"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   78
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Representantes"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   70
      Top             =   4320
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton cmdrepre 
         Caption         =   "&Nuevo"
         Height          =   450
         Index           =   0
         Left            =   10440
         Picture         =   "frmprov.frx":0D21
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdrepre 
         Caption         =   "&Modifcar"
         Height          =   450
         Index           =   1
         Left            =   10440
         Picture         =   "frmprov.frx":0E1B
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   800
         Width           =   1000
      End
      Begin VB.CommandButton cmdrepre 
         Caption         =   "&Eliminar"
         Height          =   450
         Index           =   2
         Left            =   10440
         Picture         =   "frmprov.frx":0F8D
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1900
         Width           =   1000
      End
      Begin VB.CommandButton cmdrepre 
         Caption         =   "&Grabar"
         Height          =   450
         Index           =   3
         Left            =   10440
         Picture         =   "frmprov.frx":108F
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1350
         Width           =   1000
      End
      Begin VB.CommandButton cmdrepre 
         Caption         =   "&Regresar"
         Height          =   450
         Index           =   4
         Left            =   10440
         Picture         =   "frmprov.frx":1201
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2500
         Width           =   1000
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmprov.frx":1373
         Height          =   2775
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1.5
         RowHeight       =   15
         TabAction       =   2
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
            DataField       =   "nada"
            Caption         =   "."
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
            DataField       =   "renombre"
            Caption         =   "               Nombre"
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
            DataField       =   "redireccion"
            Caption         =   "              Dirección"
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
            DataField       =   "retel1"
            Caption         =   "  Teléfono 1"
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
            DataField       =   "retel2"
            Caption         =   "Teléfono 2"
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
            DataField       =   "retipo"
            Caption         =   "Cargo"
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
            DataField       =   "reproveeedor"
            Caption         =   "Prov"
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
            MarqueeStyle    =   2
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   285.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3060.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2835.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1934.929
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   524.976
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Familias que controla el Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   48
      Top             =   360
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton Cmdagreg 
         Caption         =   "&Eliminar"
         Height          =   735
         Index           =   0
         Left            =   5400
         Picture         =   "frmprov.frx":138A
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Eliminar"
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton Cmdagreg 
         Caption         =   "&Agregar"
         Height          =   735
         Index           =   1
         Left            =   5400
         Picture         =   "frmprov.frx":17CC
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Confirmar Seleccion"
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Cmdregre 
         Caption         =   "&Regresar"
         Height          =   375
         Left            =   5040
         TabIndex        =   51
         Top             =   6240
         Width           =   1575
      End
      Begin VB.ListBox Lstlinxpro 
         DataSource      =   "Adolinxpro"
         Height          =   4155
         Left            =   6720
         TabIndex        =   50
         Top             =   960
         Width           =   4695
      End
      Begin VB.ListBox Lstlinea 
         Height          =   4155
         Left            =   240
         TabIndex        =   49
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Lblcomp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6720
         TabIndex        =   66
         Top             =   5280
         Width           =   4695
      End
      Begin VB.Label Label6 
         Caption         =   "Familias del Proveedor :"
         Height          =   255
         Left            =   6720
         TabIndex        =   53
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Catalogo de Familias :"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8000
      Left            =   120
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   11655
      Begin MSDataGridLib.DataGrid DGrprov 
         Bindings        =   "frmprov.frx":1C0E
         Height          =   6870
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   12118
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   -2147483635
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
         Caption         =   "Catálogo de proveedores"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "prove"
            Caption         =   "       Clave"
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
            DataField       =   "nomprove"
            Caption         =   "                                                 Nombre del proveedor"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               Alignment       =   2
               DividerStyle    =   0
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1514.835
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   8790.236
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   3840
         Picture         =   "frmprov.frx":1C24
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   7320
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Regresar"
         Height          =   495
         Left            =   6840
         Picture         =   "frmprov.frx":1D26
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Salir del Módulo"
         Top             =   7320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "fprov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lnuevo As Boolean
Const ColIndex_lstcar = 4
Private Sub nvalinxpro(strcvelin As String)
On Error GoTo Error:
     
     Adolinxpro.Recordset.AddNew
     Adolinxpro.Recordset!clprove = Trim(Txtprove.Text)
     Adolinxpro.Recordset!clfamilia = strcvelin
     Adolinxpro.Recordset.Update
        
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Cmdacepcom_Click()
 N = InStr(1, Lstlinea.List(Lstlinea.ListIndex), "[")
 Cvelin = Mid(Lstlinea.List(Lstlinea.ListIndex), N + 1, Len(Lstlinea.List(Lstlinea.ListIndex)) - N - 1)
 If Adolinxpro.Recordset.EOF = True And Adolinxpro.Recordset.BOF = True Then
  
     Lstlinxpro.AddItem Lstlinea.List(Lstlinea.ListIndex)
     nvalinxpro (Cvelin)
 Else
     Set RSTEMP = New ADODB.Recordset
     RSTEMP.Open "select * from linprove where clfamilia = '" & Cvelin & "' and clprove = '" & Trim(Txtprove.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     If RSTEMP.EOF Then
        Lstlinxpro.AddItem Lstlinea.List(Lstlinea.ListIndex)
        nvalinxpro (Cvelin)
     Else
        MsgBox "La linea existe con este proveedor "
     End If
     RSTEMP.Close
  End If
  Frame6.Visible = False
  Frame6.Refresh
End Sub

Private Sub adoprov_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT count( CASE activo WHEN 0 THEN CONSEC END) as proIna,count( CASE activo WHEN 1 THEN CONSEC END) as proAct FROM TFPRODUC WHERE CLAPROVE = '" & AdoProv.Recordset!prove & "' AND activo = 1", cn, adOpenStatic, adLockOptimistic, adCmdText
lblproAct.Caption = "Productos activos: " & IIf(IsNull(rs!proACT), 0, rs!proACT)
lblproina.Caption = "Productos Inactivos: " & IIf(IsNull(rs!proIna), 0, rs!proIna)
rs.Close
Set rs = Nothing
End Sub

Private Sub Cmdagreg_Click(Index As Integer)
On Error GoTo Error:
Dim RSTEMP As ADODB.Recordset
Dim N As Integer
Dim Cvelin As String
 
 Select Case Index
 Case 0
            RESPSNO = MsgBox("SE CANCELARAN TODOS LOS PRODUCTOS QUE CONTENGA ESTA RELACION " + Chr(13) + Chr(13) + "                                        DESEA CONTINUAR ? ", vbYesNo)
            If RESPSNO = vbYes Then
                N = InStr(1, Lstlinxpro.List(Lstlinxpro.ListIndex), "[")
                Cvelin = Mid(Lstlinxpro.List(Lstlinxpro.ListIndex), N + 1, Len(Lstlinxpro.List(Lstlinxpro.ListIndex)) - N - 1)
                cn.Execute " delete from linprove where clfamilia = '" & Cvelin & "' and clprove = '" & Trim(Txtprove.Text) & "'"
                Adolinxpro.Refresh
                cn.Execute "update tfproduc set actualizado = '1' , baja = '1', activo = 0, fecact=  " & date & ", fechaactivo = " & date & " where clAfamil = '" & Cvelin & "' and clAprove = '" & Trim(Txtprove.Text) & "'"
                MsgBox "TERMINÓ EL PROCESO . SI DESEA ACTIVAR NUEVAMENTE ESTOS PRODUCTOS " + Chr(13) + Chr(13) + "ASIGNARLES NUEVA FAMILIA Y PROVEEDOR EN EL CATALOGO DE PRODUCTOS "
            End If
            For i = 0 To Lstlinxpro.ListCount - 1
                If i <= (Lstlinxpro.ListCount - 1) Then
                    If Lstlinxpro.Selected(i) Then
                        Lstlinxpro.RemoveItem (i)
                        i = i - 1
                    End If
                End If
            Next i
 Case 1
             N = InStr(1, Lstlinea.List(Lstlinea.ListIndex), "[")
             Cvelin = Mid(Lstlinea.List(Lstlinea.ListIndex), N + 1, Len(Lstlinea.List(Lstlinea.ListIndex)) - N - 1)
             
             If Adolinxpro.Recordset.EOF = True And Adolinxpro.Recordset.BOF = True Then
             
                Lstlinxpro.AddItem Lstlinea.List(Lstlinea.ListIndex)
                nvalinxpro (Cvelin)
             Else
                Set RSTEMP = New ADODB.Recordset
                RSTEMP.Open "select * from linprove where clfamilia = '" & Cvelin & "' and clprove = '" & Trim(Txtprove.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
                If RSTEMP.EOF Then
                    Lstlinxpro.AddItem Lstlinea.List(Lstlinea.ListIndex)
                    nvalinxpro (Cvelin)
                Else
                    MsgBox "La linea existe con este proveedor "
                End If
                RSTEMP.Close
             End If
End Select
Exit Sub
Error:
MsgBox Err.Description
 
End Sub

Private Sub cmdcancela_Click()
On Error GoTo Error:
    AdoProv.Refresh
    Call asigna
    Call habilitar
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmddectofin_Click()
  If MsgBox("Confirma si deseas actualizar los descuentos financieros" & Chr(13) & "de todos los productos de este proveedor", vbQuestion + vbYesNo) = vbYes Then
     cn.Execute "UPDATE tfproduc SET dectofin1 = '" & txtDectoFin(0).Text & "', observa1 = '" & txtobserva(0).Text & "', dectofin2 = '" & txtDectoFin(1).Text & "', observa2 = '" & txtobserva(1).Text & "', dectofin3 = '" & txtDectoFin(2).Text & "', observa3 = '" & txtobserva(2).Text & "', dectofin4 = '" & txtDectoFin(3).Text & "', observa4 = '" & txtobserva(3).Text & "', dectofin5 = '" & txtDectoFin(4).Text & "', observa5 = '" & txtobserva(4).Text & "' WHERE claprove = '" & Trim(Txtprove.Text) & "'"
     MsgBox "La actualización de productos se realizó correctamente", vbInformation
  End If
End Sub

Private Sub Cmdlin_Click()
On Error GoTo Error:
Dim N As Integer
Dim Cvelin As String

Lstlinxpro.Clear
  Adolinxpro.CursorType = adOpenKeyset
  Adolinxpro.ConnectionString = strconnect
  Adolinxpro.CommandType = adCmdText
  Adolinxpro.RecordSource = "SELECT * FROM linprove WHERE clprove = '" & Trim(Txtprove.Text) & "'"
  Adolinxpro.Refresh
  Lstlinxpro.Clear
  
  Do While Not Adolinxpro.Recordset.EOF = True
             If Not IsNull(Adolinxpro.Recordset!clprove) Then
                For i = 0 To Lstlinea.ListCount - 1
                     N = InStr(1, Lstlinea.List(i), "[")
                     Cvelin = Mid(Lstlinea.List(i), N + 1, Len(Lstlinea.List(i)) - N - 1)
                     If Trim(Cvelin) = Trim(Adolinxpro.Recordset!clfamilia) Then
                      Lstlinxpro.AddItem Lstlinea.List(i)
                     End If
                Next
             End If
           Adolinxpro.Recordset.MoveNext
           Loop
If Adolinxpro.Recordset.EOF = False Or Adolinxpro.Recordset.BOF = False Then Adolinxpro.Recordset.MoveFirst
Frame4.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Txtprove.Visible = False
Label3.Visible = False

Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub cmdmodifica_Click()
On Error GoTo Error:
lnuevo = False
'Txtprove.SetFocus
Txtprove.Enabled = False
txtnomprove.SetFocus
Call dhabilitar
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub cmdnuevo_Click()
On Error GoTo Error:
Txtprove.Enabled = True
Txtprove.Text = ""
txtnomprove.Text = ""
txtdirpro.Text = ""
txtcolpro.Text = ""
Txtdelpro.Text = ""
Txtcodpro.Text = ""
Txttelpro.Text = ""
Txtciupro.Text = ""
Txtfrecuencia.Text = ""
Txtrfc.Text = ""
For N = 0 To 4
   txtDectoFin(N).Text = ""
   txtobserva(N).Text = ""
Next
cmbcomprador.Text = ""
Cmbtipop.Text = "": Txtfrecuencia.Text = ""
Txtrfc.Enabled = True
Cmbtipop.ListIndex = 1
lnuevo = True
Txtprove.SetFocus
Call dhabilitar
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Cmdregre_Click()
On Error GoTo Error:
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Txtprove.Visible = True
Label3.Visible = True
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub cmdrep_Click()
Adorepre.RecordSource = "select * from catrepre where reproveedor = '" & Trim(Txtprove.Text) & "'"
Adorepre.Refresh
Frame5.Visible = True
fradectos.Visible = False
Frame2.Visible = False
End Sub

Private Sub cmdrepre_Click(Index As Integer)
On Error GoTo Error
Select Case Index
Case 0
    Me.Adorepre.Recordset.AddNew
    Adorepre.Recordset!reproveedor = Me.Txtprove.Text
    Adorepre.Recordset.Update
    'Adorepre.Refresh
    Me.DataGrid1.SetFocus
Case 4
    fradectos.Visible = True
    Frame5.Visible = False
    Frame2.Visible = True
Case 2
     Adorepre.Recordset.Delete
     Adorepre.Refresh
Case 3
      Adorepre.Recordset.Update
      'Adorepre.Refresh
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdsalir_Click()
On Error GoTo Error:
Unload Me
Exit Sub
Error:
MsgBox Err.Description
End Sub


Private Sub Command1_Click()
On Error GoTo Error:
If AdoProv.Recordset.EOF = False And AdoProv.Recordset.BOF = False Then

AdoProv.Recordset.MoveFirst
Call asigna

End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command2_Click()
On Error GoTo Error:
Dim reg As Integer
reg = AdoProv.Recordset.AbsolutePosition
If reg > 1 Then
AdoProv.Recordset.MovePrevious
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command3_Click()
On Error GoTo Error:
Dim reg As Integer
Dim treg As Integer

reg = AdoProv.Recordset.AbsolutePosition
treg = AdoProv.Recordset.RecordCount
If reg < treg Then

AdoProv.Recordset.MoveNext
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command4_Click()
On Error GoTo Error:
If AdoProv.Recordset.EOF = False And AdoProv.Recordset.BOF = False Then
AdoProv.Recordset.MoveLast
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command5_Click()
'On Error GoTo error:
Dim N As Integer
Dim NUEVACLAVE As String

'VALIDACION DEL COMPRADOR
'MsgBox Len(cmbcomprador.Text)
If Len(Trim(cmbcomprador.Text)) < 1 Then
   MsgBox "Es necesario Asignar Clave del comprador", vbExclamation
   Exit Sub
End If

If Trim(txtnomprove.Text) <> "" Then
If lnuevo Then
    AdoProv.Recordset.AddNew
    NUEVACLAVE = Mid(Trim(txtnomprove.Text), 1, 1)
    Dim adorsTemp As ADODB.Recordset
    Set adorsTemp = New Recordset
    adorsTemp.LockType = adLockOptimistic
    adorsTemp.CursorType = adOpenKeyset
    adorsTemp.Source = "SELECT MAX(PROVE) AS PROVE FROM CATPROV WHERE PROVE LIKE '" & NUEVACLAVE & "%' AND PROVE <> 'BAJ' " 'AND PROVE <> 'LAC'"
    adorsTemp.ActiveConnection = cn
    adorsTemp.Open
    
    If adorsTemp.RecordCount = 1 Then
       If Val(Mid(adorsTemp!prove, 2, 2)) + 1 < 100 Then
          Txtprove.Text = NUEVACLAVE + Right("00" + Trim(Val(Mid(adorsTemp!prove, 2, 2)) + 1), 2)
       Else
          Txtprove.Text = NUEVACLAVE + Right("00" + Trim(Val(Mid(adorsTemp!prove, 2, 2)) + 1), 3)
       End If
    Else
        Txtprove.Text = NUEVACLAVE + "01"
    End If
    
    Dim NOESTA As Boolean
    Dim AUMENTA As String
    NOESTA = True
    While NOESTA
        adorsTemp.Close
        adorsTemp.Open "SELECT *  FROM CATPROV WHERE PROVE = '" & Txtprove.Text & "'", strconnect, adOpenKeyset, adLockOptimistic, adCmdText
        If adorsTemp.RecordCount < 1 Then
            NOESTA = False
        End If
        If NOESTA Then
            AUMENTAR = Val(Mid(Txtprove.Text, 2, Len(Trim(Txtprove.Text)))) + 1
            AUMENTA = Trim(Str(AUMENTAR))
        Txtprove.Text = NUEVACLAVE + IIf(Len(AUMENTA) > 1, AUMENTA, "0" & AUMENTA)
        End If
    Wend
End If

'SE GRABA LA CLAVE DE LA COMPRADORA
nposfin = Len(cmbcomprador.Text)
nPos = InStr(1, Trim(cmbcomprador.Text), "[")
nposfin = Len(Trim(cmbcomprador.Text))
'MsgBox npos
'MsgBox nposfin
Valor = nposfin - (nPos + 1)
'MsgBox valor
vcomp = Mid(cmbcomprador.Text, nPos + 1, Valor)
'MsgBox VCOMP

AdoProv.Recordset!comprador = vcomp
AdoProv.Recordset!prove = Trim(Txtprove.Text)
AdoProv.Recordset!prove = Trim(Txtprove.Text)
AdoProv.Recordset!NOMPROVE = Trim(txtnomprove.Text)
AdoProv.Recordset!dirpro = Trim(txtdirpro.Text)
AdoProv.Recordset!colpro = Trim(txtcolpro.Text)
AdoProv.Recordset!delpro = Trim(Txtdelpro.Text)
AdoProv.Recordset!codpro = Trim(Txtcodpro.Text)
AdoProv.Recordset!ciupro = Trim(Txtciupro.Text)
AdoProv.Recordset!telpro = Trim(Txttelpro.Text)
AdoProv.Recordset!frecuencia = Val(Trim(Txtfrecuencia.Text))
AdoProv.Recordset!activo = IIf(chkActivo.Value = 0, False, True)
AdoProv.Recordset!procedencia = Chktipop.Value
AdoProv.Recordset!tipo = UCase(Mid(Cmbtipop.Text, 3, 1))
AdoProv.Recordset!visita = Val(Trim(Txtvisita.Text))
AdoProv.Recordset!backorder = IIf(Chkback.Value = 0, False, True)
AdoProv.Recordset!fechaactivo = IIf(chkActivo.Value = 1, Null, date + Time)
AdoProv.Recordset!rfc = Trim(Txtrfc.Text)
AdoProv.Recordset!Volumen = chkvolumen.Value
'Igualo observaciones y descuentos financieros
For N = 1 To 5
   AdoProv.Recordset.Fields("dectofin" & CStr(N)).Value = IIf(IsNull(txtDectoFin(N - 1).Text), "", txtDectoFin(N - 1).Text)
   AdoProv.Recordset.Fields("observa" & CStr(N)).Value = IIf(IsNull(txtobserva(N - 1).Text), "", txtobserva(N - 1).Text)
Next

AdoProv.Recordset.Update
Else
    MsgBox "Favor de completar los datos ..."
    txtdesc.SetFocus
End If
Call habilitar
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command6_Click()
On Error GoTo Error:
Txtprove.Visible = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True

Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command7_Click()
On Error GoTo Error:
Frame3.Visible = False
Frame1.Visible = True
Frame2.Visible = True
Txtprove.Visible = True
Call asigna
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command8_Click()
On Error GoTo Error:
Dim siesta
Dim i As Integer
Dim cCve As String
Dim Antes

cCve = InputBox("Introduzca el nombre del proveedor a buscar", "Busqueda de Proveedor")
cCve = UCase(cCve)
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = DGrprov.Bookmark
AdoProv.Recordset.MoveFirst
DGrprov.Visible = False
For i = 1 To AdoProv.Recordset.RecordCount
AdoProv.Recordset.MoveNext
siesta = InStr(1, DGrprov.Columns(1).Text, Trim(cCve))
If siesta > 0 Then
DGrprov.Bookmark = AdoProv.Recordset.Bookmark
i = AdoProv.Recordset.RecordCount
End If
Next
DGrprov.Visible = True
If siesta = 0 Then
   MsgBox "La descripcion " & cCve & " no se encuentra en el inventario", vbExclamation
   DGrprov.Bookmark = Antes
End If
Me.DGrprov.SetFocus
Exit Sub
Error:
MsgBox Err.Description

End Sub



Private Sub Form_Deactivate()
lpprov = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
'On Error GoTo error:
Dim Strtitle As String
'CATALOGO DE COMPRADORES
Adodcusu.CommandType = adCmdText
Adodcusu.CursorType = adOpenKeyset
Adodcusu.ConnectionString = strconnect
Adodcusu.RecordSource = "select * from usuarios where caja = 777 order by NAME"
cmbcomprador.Clear
Adodcusu.Refresh

If Adodcusu.Recordset.EOF Then
 '  MsgBox "No existen Usuarios Asignados a Compras"
Else
    Adodcusu.Recordset.MoveFirst
    Do While Not Adodcusu.Recordset.EOF = True
       If Not IsNull(Adodcusu.Recordset!Name) Then
          cmbcomprador.AddItem Adodcusu.Recordset!Name + "  [" + Trim(Adodcusu.Recordset!login) + "]"
       End If
       Adodcusu.Recordset.MoveNext
    Loop
End If
Adolinea.CommandType = adCmdText
Adolinea.CursorType = adOpenKeyset
Adolinea.ConnectionString = strconnect
Adolinea.RecordSource = "select * from familias order by fdescrip"

Lstlinea.Clear
Adolinea.Refresh
    Do While Not Adolinea.Recordset.EOF = True
               If Not IsNull(Adolinea.Recordset!fdescrip) Then
                 Lstlinea.AddItem Adolinea.Recordset!fdescrip + "  [" + Adolinea.Recordset!fclave + "]"
               End If
              Adolinea.Recordset.MoveNext
            Loop
           Adolinea.Recordset.MoveFirst


AdoProv.CursorType = adOpenKeyset
AdoProv.CommandType = adCmdText
AdoProv.ConnectionString = cn.ConnectionString
AdoProv.RecordSource = "SELECT * FROM CATPROV ORDER BY NOMPROVE "
AdoProv.Refresh

Adorepre.CommandType = adCmdText
Adorepre.CursorType = adOpenKeyset
Adorepre.LockType = adLockOptimistic

'Adorepre.ConnectionString = cn.ConnectionString
Adorepre.ConnectionString = cCadConex



Strtitle = Str(AdoProv.Recordset.RecordCount)
fprov.Caption = "Catalogo de Proveedores...   Total :" + Strtitle
Call asigna
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub asigna()
Dim RSTEMP As ADODB.Recordset
Dim nclavep As String
'On Error GoTo error:
If AdoProv.Recordset.EOF = False And AdoProv.Recordset.BOF = False Then
If lpprov Then
    AdoProv.Recordset.Find "prove = '" & strcveprov & "'"
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = False
    cmdnuevo.Enabled = False
End If

Txtprove.Text = AdoProv.Recordset!prove
txtnomprove.Text = IIf(Not IsNull(AdoProv.Recordset!NOMPROVE), AdoProv.Recordset!NOMPROVE, "")
txtdirpro.Text = IIf(Not IsNull(AdoProv.Recordset!dirpro), AdoProv.Recordset!dirpro, "")
txtcolpro.Text = IIf(Not IsNull(AdoProv.Recordset!colpro), AdoProv.Recordset!colpro, "")
Txtdelpro.Text = IIf(Not IsNull(AdoProv.Recordset!delpro), AdoProv.Recordset!delpro, "")
Txtcodpro.Text = IIf(Not IsNull(AdoProv.Recordset!codpro), AdoProv.Recordset!codpro, "")
Txttelpro.Text = IIf(Not IsNull(AdoProv.Recordset!telpro), AdoProv.Recordset!telpro, "")
Txtciupro.Text = IIf(Not IsNull(AdoProv.Recordset!ciupro), AdoProv.Recordset!ciupro, "")
Txtvisita.Text = IIf(Not IsNull(AdoProv.Recordset!visita), AdoProv.Recordset!visita, "")
Txtfrecuencia.Text = IIf(Not IsNull(AdoProv.Recordset!frecuencia), AdoProv.Recordset!frecuencia, "")
Txtrfc.Text = IIf(Not IsNull(AdoProv.Recordset!rfc), AdoProv.Recordset!rfc, "")
chkActivo.Value = IIf(AdoProv.Recordset!activo, 1, 0)
Chkback.Value = IIf(AdoProv.Recordset!backorder, 1, 0)
Chktipop.Value = IIf(Not IsNull(AdoProv.Recordset!procedencia), AdoProv.Recordset!procedencia, 0)
lblactivo = "Fecha Baja : " & IIf(IsNull(AdoProv.Recordset!fechaactivo), " ", AdoProv.Recordset!fechaactivo)
'Igualo observaciones y descuentos financieros
For N = 1 To 5
   txtDectoFin(N - 1).Text = IIf(IsNull(AdoProv.Recordset.Fields("dectofin" & CStr(N)).Value), "", AdoProv.Recordset.Fields("dectofin" & CStr(N)).Value)
   txtobserva(N - 1).Text = IIf(IsNull(AdoProv.Recordset.Fields("observa" & CStr(N)).Value), "", AdoProv.Recordset.Fields("observa" & CStr(N)).Value)
Next


If Len(Trim(lblactivo.Caption)) > 13 Then
    lblactivo.Visible = True
    lblactivo.Refresh
Else
    lblactivo.Visible = False
    lblactivo.Refresh
End If

'SE PONE LA CLAVE DEL COMPRADOR SI ES QUE LO TIENE REGISTRADO
Me.cmbcomprador.Text = " "
'CBMCOMPRADOR.Text = " "
vcomp = AdoProv.Recordset!comprador
'SE BUSCA EN LA TABLA DE USUARIOS
Set RSTEMP = New ADODB.Recordset
     RSTEMP.Open "select * from usuarios  where caja  = 777  and login =  '" & Trim(vcomp) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     If RSTEMP.EOF Then
        'MsgBox "error en usuario"
    Else
       cmbcomprador.Text = RSTEMP!Name + "  [" + Trim(RSTEMP!login) + "]"
    End If
RSTEMP.Close
Cmbtipop.Text = ""
For i = 0 To 3
    If Cmbtipop.Text = "" Then
    A = UCase(Mid(Cmbtipop.List(i), 3, 1))
    If UCase(Mid(Cmbtipop.List(i), 3, 1)) = AdoProv.Recordset!tipo Then
     Cmbtipop.Text = Cmbtipop.List(i)
     Cmbtipop.ListIndex = i
    End If
    End If
Next
If Cmbtipop.Text = "" Then
  Cmbtipop.Text = "Indirecto"
  Cmbtipop.ListIndex = 0
End If
nclavep = IIf(Not IsNull(AdoProv.Recordset!comprador), AdoProv.Recordset!comprador, "000")

lnuevo = False
 
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub


Private Sub dhabilitar()
On Error GoTo Error:
 Command5.Enabled = True
 cmdcancela.Enabled = True
 cmdnuevo.Enabled = False
 cmdmodifica.Enabled = False
 Command1.Enabled = False
 Command2.Enabled = False
 Command3.Enabled = False
 Command4.Enabled = False
 Command6.Enabled = False
Exit Sub
Error:
MsgBox Err.Description
End Sub
Private Sub habilitar()
On Error GoTo Error:
 Command5.Enabled = False
 cmdcancela.Enabled = False
 cmdmodifica.Enabled = True
  If lpprov = False Then
 'cmdnuevo.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 Command3.Enabled = True
 Command4.Enabled = True
 Command6.Enabled = True
 End If
Exit Sub
Error:
MsgBox Err.Description
 End Sub

Private Sub LstCargo_DblClick()
 DGrrepre.Columns(4).Text = LstCargo.List(LstCargo.ListIndex)
 LstCargo.Visible = False
End Sub

Private Sub LstCargo_LostFocus()
DGrrepre.Columns(4).Text = LstCargo.List(LstCargo.ListIndex)
 LstCargo.Visible = False
End Sub



Private Sub Txtfrecuencia_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub



Private Sub txtnomprove_LostFocus()
  txtnomprove.Text = UCase(txtnomprove.Text)
  txtnomprove.Refresh
End Sub

Private Sub Txtprove_LostFocus()
     Txtprove.Text = UCase(Txtprove.Text)
     Txtprove.Refresh
     If Len(Trim(Txtprove.Text)) > 0 Then
        AdoProv.Refresh
        AdoProv.Recordset.MoveFirst
        AdoProv.Recordset.Find "prove = '" & UCase(Trim(Txtprove.Text)) & "'"
        If AdoProv.Recordset.EOF = False Then
           Call asigna
        End If
      End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtprove_Validate(Cancel As Boolean)
Txtprove_LostFocus
End Sub


Private Sub Txtvisita_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub
