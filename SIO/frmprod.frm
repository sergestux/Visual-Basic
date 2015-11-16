VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmprod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Especificaciones de productos"
   ClientHeight    =   6510
   ClientLeft      =   4155
   ClientTop       =   1545
   ClientWidth     =   9540
   Icon            =   "frmprod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9540
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtpeso 
      Alignment       =   2  'Center
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
      Left            =   8400
      MaxLength       =   25
      TabIndex        =   22
      Top             =   5040
      Width           =   855
   End
   Begin ComctlLib.ProgressBar probar1 
      Height          =   255
      Left            =   120
      TabIndex        =   76
      Top             =   7560
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc adolineas 
      Height          =   330
      Left            =   120
      Top             =   6240
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
      Caption         =   "lineas"
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
   Begin MSAdodcLib.Adodc Adoprecio 
      Height          =   330
      Left            =   2040
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
   Begin VB.Frame frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   50
      Top             =   5520
      Width           =   9255
      Begin VB.CommandButton Cmdact 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5880
         Picture         =   "frmprod.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Guardar los cambios"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmdeof 
         Height          =   495
         Left            =   2040
         Picture         =   "frmprod.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ir al ultimo registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmdnext 
         Height          =   495
         Left            =   1440
         Picture         =   "frmprod.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ir al siguiente registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmdback 
         Height          =   495
         Left            =   840
         Picture         =   "frmprod.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ir al registro anterior"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmdbof 
         Height          =   495
         Left            =   240
         Picture         =   "frmprod.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ir al primer registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdcancela 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6840
         Picture         =   "frmprod.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cancelar la Captura"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Regresar"
         Height          =   495
         Left            =   7800
         Picture         =   "frmprod.frx":0BB6
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Salir del modulo"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdmodifica 
         Caption         =   "&Modificar"
         Height          =   495
         Left            =   3960
         Picture         =   "frmprod.frx":0D28
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Modificar datos del producto"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4920
         Picture         =   "frmprod.frx":0E9A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Nuevo Producto"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmdbus 
         Height          =   495
         Left            =   2640
         Picture         =   "frmprod.frx":0F94
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Buscar Registro"
         Top             =   240
         Width           =   495
      End
      Begin MSAdodcLib.Adodc Adocatprod 
         Height          =   375
         Left            =   -1680
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
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
         ConnectStringType=   3
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
         Caption         =   "productos"
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
      Begin MSAdodcLib.Adodc Adoprov 
         Height          =   330
         Left            =   5040
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSAdodcLib.Adodc Adofamilia 
         Height          =   330
         Left            =   6480
         Top             =   0
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   38
      Top             =   4440
      Width           =   9255
      Begin VB.TextBox txtvolumen 
         Alignment       =   2  'Center
         DataField       =   "volmtocub"
         DataSource      =   "Adocatprod"
         Enabled         =   0   'False
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
         Left            =   7200
         MaxLength       =   25
         TabIndex        =   81
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtlargo 
         Alignment       =   2  'Center
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
         Left            =   4560
         MaxLength       =   25
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtalto 
         Alignment       =   2  'Center
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
         Left            =   6240
         MaxLength       =   25
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtancho 
         Alignment       =   2  'Center
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
         Left            =   5400
         MaxLength       =   25
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Txtpzaxca 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Txtprecosca 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   15
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Txtcoprecio 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mts.cúbicos Volúmen  "
         Height          =   495
         Left            =   7200
         TabIndex        =   82
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Kilos"
         Height          =   255
         Left            =   8280
         TabIndex        =   80
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Centimetros   lineales"
         Height          =   255
         Left            =   4560
         TabIndex        =   79
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Alto"
         Height          =   255
         Left            =   6240
         TabIndex        =   78
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Peso"
         Height          =   255
         Left            =   8280
         TabIndex        =   77
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   60
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Pzas/Caja"
         Height          =   255
         Left            =   1800
         TabIndex        =   58
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Precio Lista"
         Height          =   255
         Left            =   360
         TabIndex        =   57
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho"
         Height          =   255
         Left            =   5400
         TabIndex        =   49
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl11 
         Alignment       =   2  'Center
         Caption         =   "Largo"
         Height          =   255
         Left            =   4560
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Costo / Unidad"
         Height          =   255
         Left            =   2880
         TabIndex        =   47
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Caracteristicas de Producto"
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
      Height          =   4455
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   9255
      Begin VB.CheckBox Chkcosituacion 
         Caption         =   "Activo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkProm 
         Caption         =   "Promoción especial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Txtbarrascaja 
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
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cmbtipo 
         Height          =   315
         ItemData        =   "frmprod.frx":108E
         Left            =   6480
         List            =   "frmprod.frx":10A1
         TabIndex        =   15
         Top             =   4035
         Width           =   2655
      End
      Begin VB.TextBox Txtclavdprov 
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
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox Chknal 
         Caption         =   "Nacional"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   7080
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Txtnomcor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   7
         Top             =   1560
         Width           =   7215
      End
      Begin VB.TextBox Txtclafamil 
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
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   13
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Txtclaprove 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   11
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ComboBox Cmbmedida 
         DataField       =   "medida"
         DataSource      =   "Adocatprod"
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
         ItemData        =   "frmprod.frx":10ED
         Left            =   3120
         List            =   "frmprod.frx":1106
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox Cmbfamilia 
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
         Left            =   2640
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   3000
         Width           =   6375
      End
      Begin VB.ComboBox Cmbprove 
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
         ItemData        =   "frmprod.frx":1131
         Left            =   2640
         List            =   "frmprod.frx":1133
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   2520
         Width           =   6375
      End
      Begin VB.TextBox Txtbarraspza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Txtcontenido 
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
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Txtdescripc 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1080
         Width           =   7215
      End
      Begin VB.TextBox Txtconsec 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbusuario 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "zzz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   3840
         TabIndex        =   75
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label24 
         Caption         =   "Codigo Caja"
         Height          =   255
         Left            =   6000
         TabIndex        =   74
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label fechaact 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "zzz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1920
         TabIndex        =   73
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label23 
         Caption         =   "Fecha Actualizacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Producto:"
         Height          =   255
         Left            =   6480
         TabIndex        =   71
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label lblcprov 
         Caption         =   "Clave del Proveedor:"
         Height          =   255
         Left            =   4800
         TabIndex        =   69
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Corto :"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Lbldepto 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   55
         Top             =   3720
         Width           =   7455
      End
      Begin VB.Label Lbfam 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1320
         TabIndex        =   54
         Top             =   3360
         Width           =   7695
      End
      Begin VB.Label Label14 
         Caption         =   "Medida"
         Height          =   255
         Left            =   2520
         TabIndex        =   53
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento:"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Linea :"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Familia :"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Proveedor:"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo de Barras"
         Height          =   375
         Left            =   2520
         TabIndex        =   42
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Contenido"
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre del Producto"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Catalogo de Productos"
      Height          =   6495
      Left            =   120
      TabIndex        =   51
      Top             =   480
      Visible         =   0   'False
      Width           =   9255
      Begin MSDataGridLib.DataGrid DGrprod 
         Bindings        =   "frmprod.frx":1135
         Height          =   5295
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
            DataField       =   "claprove"
            Caption         =   "Prov."
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
         BeginProperty Column03 
            DataField       =   "contenid"
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
         BeginProperty Column04 
            DataField       =   "medida"
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
         BeginProperty Column05 
            DataField       =   "barraspza"
            Caption         =   "Cod.Barras"
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
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   4784.882
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1725.165
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Cmdbusqueda 
         Caption         =   "Codigo &Barras"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   68
         ToolTipText     =   "Busqueda por Clave del Producto"
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton Cmdbusqueda 
         Caption         =   "&Clave"
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   65
         ToolTipText     =   "Busqueda por Clave del Producto"
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton Cmdbusqueda 
         Caption         =   "&Descripcion"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   64
         ToolTipText     =   "Busqueda por Descripcion del Producto"
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Regresar"
         Height          =   375
         Left            =   5640
         TabIndex        =   52
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Presentacion"
         Height          =   255
         Left            =   6960
         TabIndex        =   63
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Precios"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   66
      Top             =   6600
      Visible         =   0   'False
      Width           =   9015
      Begin VB.TextBox Txtprecio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtprecio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtprecio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtprecio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "       Menudeo                   May. Autoservicio            May. Vendedores               May. Bodega"
         Height          =   255
         Left            =   360
         TabIndex        =   67
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Label Lblconsulta 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando Informacion ...... "
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
      Left            =   120
      TabIndex        =   70
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label13 
      Height          =   375
      Left            =   5880
      TabIndex        =   56
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frmprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lnuevo As Boolean
Dim lsibusca As Boolean

Private Sub Cmbfamilia_LostFocus()
'On Error GoTo error:
Dim strfam As String
Dim N As Integer
            If Trim(Cmbfamilia.Text) <> "" And Cmbfamilia.ListIndex > -1 Then
             N = InStr(1, Cmbfamilia.List(Cmbfamilia.ListIndex), "[")
             strfam = Mid(Cmbfamilia.List(Cmbfamilia.ListIndex), N + 1, Len(Cmbfamilia.List(Cmbfamilia.ListIndex)) - N - 1)
             Adofamilia.Recordset.MoveFirst
             Adofamilia.Recordset.Find "sfclave = '" & strfam & "'"
             If Adofamilia.Recordset.EOF Then
                MsgBox "Seleccione una linea "
                Cmbfamilia.SetFocus
             Else
                         
             Txtclafamil.Text = Adofamilia.Recordset!sfclave
             Lbfam.Caption = Adofamilia.Recordset!fdescrip
             Lbldepto.Caption = Adofamilia.Recordset!depdescrip
             'Txtcoprecio.SetFocus
             End If
             Else
             If Trim(Cmbfamilia.Text) = "" Then
                MsgBox "Seleccione una linea"
                'Cmbfamilia.SetFocus
             End If
             End If

Exit Sub
Error:
MsgBox Err.Description
End Sub



Private Sub Cmbprove_LostFocus()
'On Error GoTo error:
Dim rs As ADODB.Recordset
Dim strprove As String
Dim N As Integer
If lsibusca Then
    If Trim(Cmbprove.Text) <> "" And Cmbprove.ListIndex > -1 Then
            
             N = InStr(1, Cmbprove.List(Cmbprove.ListIndex), "[")
             strprove = Mid(Cmbprove.List(Cmbprove.ListIndex), N + 1, Len(Cmbprove.List(Cmbprove.ListIndex)) - N - 1)
              Txtclaprove.Text = strprove
             Set rs = New ADODB.Recordset
                rs.LockType = adLockOptimistic
                rs.CursorType = adOpenKeyset
                rs.ActiveConnection = cn
                
             'cad = "select * from linprove,familias,lineas,departamento " & _
             '      "WHERE clprove = '" & Trim(Txtclaprove.Text) & "' and clfamilia = fclave  and sffamilia = fclave AND fdepto = depclave "
             CAD = "select * from familias,lineas,departamento " & _
                   "WHERE sffamilia = fclave AND fdepto = depclave "
             rs.Source = CAD
             rs.Open
             Set Adofamilia.Recordset = rs
             Adofamilia.Refresh
             Call llenafamilia
             Txtclafamil.SetFocus
             Else
             If Trim(Cmbprove.Text) = "" Then
                MsgBox "Seleccione un proveedor"
                Cmbprove.SetFocus
             End If
      End If
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Cmdact_Click()
Dim rs As ADODB.Recordset
On Error GoTo Error:
If Trim(Txtdescripc.Text) <> "" Then
    If lnuevo Then
       'Validar que el código de barras no este asignado algun otro producto
       Set rs = New ADODB.Recordset
       rs.Open "SELECT * FROM TFPRODUC WHERE BARRASPZA = " & IIf(Txtbarraspza.Text = "", 0, Txtbarraspza.Text), cn, adOpenDynamic, adLockOptimistic, adCmdText
       If (rs.BOF And rs.EOF) Or Trim(Txtbarraspza.Text) = "0" Or Trim(Txtbarraspza.Text) = "" Then
          Adocatprod.Refresh
          Adocatprod.Recordset.AddNew
       Else
          MsgBox "EL CODIGO DE BARRAS " & Txtbarraspza.Text & " YA ESTA ASIGNADO AL PRODUCTO " & rs!DESCRIPC & "  " & Str(rs!PAQUETES) & " x " & CStr(rs!Contenid) & " " + rs!medida, vbExclamation
          Exit Sub
       End If
    End If
    
    If Len(Trim(Txtclaprove.Text)) > 0 Then
       Adocatprod.Recordset!claprove = Trim(Txtclaprove.Text)
    Else
       MsgBox "DEBE ASIGNAR UN PROVEEDOR", vbExclamation
    End If
    'VALIDACION DE LA FAMILIA
    If Len(Trim(Txtclafamil.Text)) > 0 Then
        Adocatprod.Recordset!linea = Trim(Txtclafamil.Text)
    Else
        MsgBox "DEBE ASIGNAR UNA FAMILIA ...", vbExclamation
        'Exit Sub
    End If
    Adocatprod.Recordset!INTERNO = IIf(Nivel = "I", 1, 0)
    Adocatprod.Recordset!DESCRIPC = Trim(Txtdescripc.Text)
    Adocatprod.Recordset!Contenid = IIf(Val(Txtcontenido.Text) = 0, 1, Val(Txtcontenido.Text))
    Adocatprod.Recordset!barraspza = IIf(Trim(Txtbarraspza.Text) <> "", Trim(Txtbarraspza.Text), 0)
    Adocatprod.Recordset!barrascaja = IIf(Trim(Txtbarrascaja.Text) <> "", Trim(Txtbarrascaja.Text), 0)
    Adocatprod.Recordset!medida = Trim(Cmbmedida.Text)
    Adocatprod.Recordset!NOMCORTO = Trim(Txtnomcor.Text)
    Adocatprod.Recordset!procedencia = Chknal.Value
    Adocatprod.Recordset!activo = Chkcosituacion.Value
    Adocatprod.Recordset!enpromo = chkProm.Value
    Adocatprod.Recordset!fechaactivo = IIf(Chkcosituacion.Value = 0, date, Null)
    Adocatprod.Recordset!baja = IIf(Chkcosituacion.Value = 0, "1", "0")
    Adocatprod.Recordset!actualizado = "1"
    Adocatprod.Recordset!fecact = date
    Adocatprod.Recordset!peso = Trim(txtpeso.Text)
    Adocatprod.Recordset!largo = Trim(txtlargo.Text)
    Adocatprod.Recordset!ancho = Trim(txtancho.Text)
    Adocatprod.Recordset!Alto = Trim(txtalto.Text)
    Adocatprod.Recordset!Volmtocub = IIf((Trim(txtvolumen.Text) = ""), 0, Trim(txtvolumen.Text))
    'Adocatprod.Recordset!clafamil = Trim(txtfamilant.Text)
    Adocatprod.Recordset!clavedelprov = Trim(Txtclavdprov.Text)
    If Adofamilia.Recordset.EOF = False Or Adofamilia.Recordset.BOF = False Then Adocatprod.Recordset!clafamil = Adofamilia.Recordset!fclave
    Adocatprod.Recordset!PAQUETES = IIf(Val(Txtpzaxca.Text) = 0, 1, Trim(Txtpzaxca.Text))
    Adocatprod.Recordset!costocaj = IIf(Val(Txtprecosca.Text) = 0, 1, Trim(Txtprecosca.Text))
    Adocatprod.Recordset!tipoproducto = Trim(cmbtipo.Text)
    Adocatprod.Recordset!USUARIO = Trim(cUsuario)
    If lnuevo Then
        generamaximo
        Adocatprod.Recordset!CONSEC = Trim(Txtconsec.Text)
        'Adocatprod.Recordset!CONSEC = "7000000"
    End If
    Adocatprod.Recordset.Update
    fechaact.Caption = date
    lbusuario.Caption = "MODIFICO  : " & Trim(cUsuario)
    Call habilitar
    Unload Me
     If lnuevo Then
        'PRESENTAR PANTALLA DE PRECIOS Y DESCUENTOS
        strcveprod = Adocatprod.Recordset!CONSEC
        lpprov = True
        lpprod = False
        If tipotienda <> 3 Then
           frmprecios.Show 1
        Else
           fnewprec.Show 1
        End If
    End If
Else
    MsgBox " ¡¡  No se puede grabar un producto sin NOMBRE !!", vbCritical
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub CmdBack_Click()

'On Error GoTo error:

Dim reg As Integer
reg = Adocatprod.Recordset.Bookmark
If reg > 1 Then
    Adocatprod.Recordset.MovePrevious
Call asigna
End If

Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Cmdbof_Click()
'On Error GoTo error:
    If Adocatprod.Recordset.EOF = False And Adocatprod.Recordset.BOF = False Then
        Adocatprod.Recordset.MoveFirst
        Call asigna
    End If

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Cmdbus_Click()
'On Error GoTo error:
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame5.Visible = False
Frame4.Visible = True
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Cmdbusqueda_Click(Index As Integer)
'On Error GoTo error:
DGrprod.Visible = False
siesta = 0
Dim cCve As String
Dim Antes

Select Case Index
Case 0
cCve = InputBox("Introduzca el NOMBRE del producto a buscar", "Busqueda de Producto")
Case 1
cCve = InputBox("Introduzca el CODIGO DE BARRAS del producto ", "Busqueda de Producto")
Case 2
cCve = InputBox("Introduzca la CLAVE del producto a buscar", "Busqueda de Producto")
End Select

If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub

cCve = UCase(cCve)

Antes = DGrprod.Bookmark


Select Case Index
Case 0

Adocatprod.Recordset.MoveFirst
Adocatprod.Recordset.Find "descripc like '" & Trim(cCve) & "*'"


Case 1

Adocatprod.Recordset.MoveFirst
Adocatprod.Recordset.Find "barraspza = " & Trim(cCve)
Case 2


Adocatprod.Recordset.MoveFirst

Adocatprod.Recordset.Find "consec like '" & Trim(cCve) & "*'"

End Select


If Adocatprod.Recordset.EOF Then
    Select Case Index
    Case 0
        MsgBox "La DESCRIPCION " & cCve & " no se encuentra en el Catalogo", vbExclamation
    Case 1
        MsgBox "El CODIGO DE BARRAS " & cCve & " no se encuentra en el Catalogo", vbExclamation
    Case 2
        MsgBox "La CLAVE " & cCve & " no se encuentra en el Catalogo", vbExclamation
    End Select
   DGrprod.Bookmark = Antes
End If

DGrprod.Visible = True

Exit Sub
Error:
MsgBox Err.Description


End Sub


Private Sub cmdcancela_Click()
'On Error GoTo error:
Call habilitar
Call asigna
lnuevo = False

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Cmdeof_Click()
'On Error GoTo error:
If Adocatprod.Recordset.EOF = False And Adocatprod.Recordset.BOF = False Then
Adocatprod.Recordset.MoveLast
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub


Private Sub cmdmodifica_Click()
'On Error GoTo error:
Txtbarraspza.SetFocus
Call dhabilitar
lnuevo = False
lsibusca = True
If tipotienda = 2 Then
   Me.Frame1.Enabled = False
   Frame2.Enabled = False
   'txtpeso.Enabled = True
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
'On Error GoTo error:
Dim preg As Integer
Dim ureg As Integer
preg = Adocatprod.Recordset.Bookmark
ureg = Adocatprod.Recordset.RecordCount
If preg < ureg Then
    Adocatprod.Recordset.MoveNext
    Call asigna
End If

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdnuevo_Click()
'On Error GoTo error:
lsibusca = True
lnuevo = True
Call nuevoprod
Call dhabilitar
Txtbarraspza.SetFocus
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdsalir_Click()
'On Error GoTo error:
Unload Me
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
'On Error GoTo error:
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = True
Frame5.Visible = True
Frame4.Visible = False
Cmdeof.SetFocus
Call asigna
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command2_Click()
'On Error GoTo error:
Unload Me
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'On Error GoTo error:
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
ElseIf KeyAscii = 27 Then
   Unload Me
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
'On Error GoTo error:
Dim rs As ADODB.Recordset
Dim strcont As String
If lpprod = False Then
strcont = "select * from tfproduc order by descripc"
Else
strcont = "select * from tfproduc where consec = '" & Trim(strcveprod) & "'  order by descripc "
End If

Adocatprod.CommandType = adCmdText
Adocatprod.CursorType = adOpenKeyset
Adocatprod.RecordSource = strcont
Adocatprod.ConnectionString = strconnect
Adocatprod.Refresh

AdoProv.CommandType = adCmdText
AdoProv.CursorType = adOpenKeyset
AdoProv.ConnectionString = strconnect
AdoProv.RecordSource = "select * from catprov order by nomprove"
AdoProv.Refresh

Cmbprove.Clear
Do While Not AdoProv.Recordset.EOF
             If Not IsNull(AdoProv.Recordset!NOMPROVE) Then
             Cmbprove.AddItem AdoProv.Recordset!NOMPROVE + "  Clave [" + AdoProv.Recordset!prove + "]"
            End If
             AdoProv.Recordset.MoveNext
 Loop
AdoProv.Recordset.MoveFirst

    
Adofamilia.CommandType = adCmdText
Adofamilia.CursorType = adOpenKeyset
Adofamilia.ConnectionString = strconnect
CAD = "select * from departamento,familias,lineas WHERE fdepto = depclave and fclave = sffamilia"
Adofamilia.RecordSource = CAD
Adofamilia.Refresh
Call llenafamilia

'SI SE ENTRA POR HOJA DE CATALOGO
If lpprod Then
If Adocatprod.Recordset.RecordCount > 0 Then
    Adocatprod.Recordset.Find " consec = ' " & Trim(strcveprod) & "'"
End If
End If
Call asigna

Me.cmdmodifica.Enabled = (Nivel <> "P") And tipotienda <> 3
'Me.cmdnuevo.Enabled = (Nivel <> "P") And tipotienda <> 3
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
 consulta = ""
End Sub

Private Sub txtalto_LostFocus()
txtvolumen.Text = Round(txtlargo.Text * txtancho.Text * txtalto.Text / 1000000, 4)
End Sub

Private Sub txtancho_LostFocus()
txtvolumen.Text = Round(txtlargo.Text * txtancho.Text * txtalto.Text / 1000000, 4)
End Sub

Private Sub Txtbarraspza_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If
End Sub


Private Sub Txtclavdprov_LostFocus()
Txtclavdprov.Text = UCase(Txtclavdprov.Text)
End Sub


Private Sub Txtcontenido_KeyPress(KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If
End Sub

Private Sub Txtcoprecio_KeyPress(KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub

Private Sub txtlargo_LostFocus()
txtvolumen.Text = Round(txtlargo.Text * txtancho.Text * txtalto.Text / 1000000, 4)
End Sub

Private Sub Txtprecosca_KeyPress(KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub

Private Sub Txtprecosca_LostFocus()

'On Error GoTo error:
If Trim(Txtpzaxca.Text) <> "" And Trim(Txtprecosca.Text) <> "" Then
If Val(Txtpzaxca.Text) <> 0 And Val(Txtprecosca.Text) <> 0 Then
Txtcoprecio.Text = Round(Val(Txtprecosca.Text) / Val(Txtpzaxca.Text), 2)
Else
Txtcoprecio.Text = Txtprecosca.Text
End If
End If

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtbarraspza_LostFocus()
Dim reg
'On Error GoTo error:
' se checa si el codigo ya existe
'If Trim(Txtbarraspza.Text) <> "" Then
'           Set rstTemp = New ADODB.Recordset
'           rstTemp.ActiveConnection = cn
'           rstTemp.CursorType = adOpenKeyset
'           rstTemp.Source = "SELECT * from  [tfproduc] WHERE barraspza  = '" & Trim(Me.Txtbarraspza.Text) & "'"
'           rstTemp.Open
'           If Not rstTemp.EOF Then
'              MsgBox "ESTE CODIGO YA EXISTE, EN LA CLAVE " & rstTemp!CONSEC, vbInformation
'          End If
'          rstTemp.Close
'reg = Adocatprod.Recordset.Bookmark
'Adocatprod.Recordset.MoveFirst
'Adocatprod.Recordset.Find "barraspza = " & Trim(Txtbarraspza.Text) & ""
'If Adocatprod.Recordset.EOF = False Then
'    If Trim(Adocatprod.Recordset!CONSEC) <> Trim(Txtconsec.Text) And Trim(Txtbarraspza.Text) <> "0" And Cmdact.Enabled = True Then
'    MsgBox "Este codigo de Barras ya existe "
'    Txtbarraspza.Text = ""
'    Txtbarraspza.SetFocus
'    End If
'    If Cmdact.Enabled = False Then
'        Call asigna
'    End If
'Else
'    If Cmdact.Enabled = False Then
'    MsgBox "Este codigo de Barras no existe "
'    End If
'End If
'Adocatprod.Recordset.Bookmark = reg
'End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtclafamil_LostFocus()
'On Error GoTo error:
            Txtclafamil.Text = UCase(Txtclafamil.Text)
            If Trim(Txtclafamil.Text) <> "" Then
             Adofamilia.Refresh
             If Not (Adofamilia.Recordset.BOF And Adofamilia.Recordset.EOF) Then Adofamilia.Recordset.MoveFirst
             Adofamilia.Recordset.Find "sfclave = '" & Trim(Txtclafamil.Text) & "'"
             If Adofamilia.Recordset.EOF Then
                MsgBox "Seleccione la linea "
                Cmbfamilia.SetFocus
             Else
                Cmbfamilia.Text = Adofamilia.Recordset!sfdescrip + "[" + Adofamilia.Recordset!sfclave + "]"
                Txtclafamil.Text = Adofamilia.Recordset!sfclave
                Lbfam.Caption = Adofamilia.Recordset!fdescrip
                Lbldepto.Caption = Adofamilia.Recordset!depdescrip
             End If
             Else
             'MsgBox "Seleccione la  Linea  "
             Cmbfamilia.SetFocus
             
End If

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtclaprove_LostFocus()
Dim lsiprov As Boolean
Dim RSTEMP As ADODB.Recordset
Set RSTEMP = New ADODB.Recordset
'On Error GoTo error:
Txtclaprove.Text = UCase(Txtclaprove.Text)
Txtclaprove.Refresh
Cmbprove.Locked = True

If Trim(Txtclaprove.Text) <> "" Then
       lsiprov = True
       For i = 0 To Cmbprove.ListCount - 1
              Cmbprove.ListIndex = i
              N = InStr(1, Cmbprove.List(Cmbprove.ListIndex), "[")
             strprove = Mid(Cmbprove.List(Cmbprove.ListIndex), N + 1, Len(Cmbprove.List(Cmbprove.ListIndex)) - N - 1)
             lsibusca = False
                      
          If Trim(Txtclaprove.Text) = Trim(strprove) Then
                Cmbprove.ListIndex = i
                
                'Adofamilia.CommandType = adCmdText
                'Adofamilia.CursorType = adOpenKeyset
                'Adofamilia.ConnectionString = strconnect
                'cad = "select * from familias,lineas,departamento,linprove where fclave = clfamilia and sffamilia = fclave AND fdepto = depclave AND clprove = '" & Trim(Txtclaprove.Text) & "'"
                'Adofamilia.RecordSource = cad
                'Adofamilia.Refresh
                
                'Call llenafamilia
                lsiprov = False
                Cmbfamilia.SetFocus
                i = Cmbprove.ListCount
          End If
      Next
      If lsiprov Then
            MsgBox "Seleccione un proveedor "
            Cmbprove.SetFocus
            lsibusca = True
      End If
Else
       MsgBox "Seleccione un proveedor "
       Cmbprove.SetFocus
End If
Cmbprove.Locked = False

Exit Sub
Error:
MsgBox Err.Description

End Sub


Private Sub asigna()
'On Error GoTo error:
Dim RSTEMP As ADODB.Recordset
Set RSTEMP = New ADODB.Recordset
Dim strprov As String
If lpprod Then
    Adocatprod.Recordset.MoveFirst
    Adocatprod.Recordset.Find "consec = '" & Trim(strcveprod) & "'"
    lpprod = False
End If

If Adocatprod.Recordset.EOF = False And Adocatprod.Recordset.BOF = False Then
    Txtconsec.Text = IIf(Not IsNull(Adocatprod.Recordset!CONSEC), Adocatprod.Recordset!CONSEC, "")
    Txtdescripc.Text = IIf(Not IsNull(Adocatprod.Recordset!DESCRIPC), Adocatprod.Recordset!DESCRIPC, "")
    Txtcontenido.Text = IIf(Not IsNull(Adocatprod.Recordset!Contenid), Adocatprod.Recordset!Contenid, "")
    Txtbarraspza.Text = IIf(Not IsNull(Adocatprod.Recordset!barraspza), Adocatprod.Recordset!barraspza, "")
    Txtbarrascaja.Text = IIf(Not IsNull(Adocatprod.Recordset!barrascaja), Adocatprod.Recordset!barrascaja, "")
'    Cmbmedida.Text = IIf(Not IsNull(Adocatprod.Recordset!medida), Adocatprod.Recordset!medida, "")
    Txtclafamil.Text = IIf(Not IsNull(Adocatprod.Recordset!linea), Adocatprod.Recordset!linea, "")
    Txtclaprove.Text = IIf(Not IsNull(Adocatprod.Recordset!claprove), Trim(Adocatprod.Recordset!claprove), "")
    
    Txtprecosca.Text = IIf(Not IsNull(Adocatprod.Recordset!costocaj), Adocatprod.Recordset!costocaj, "")
    Txtpzaxca.Text = IIf(Not IsNull(Adocatprod.Recordset!PAQUETES), Adocatprod.Recordset!PAQUETES, "")
    Txtnomcor.Text = IIf(Not IsNull(Adocatprod.Recordset!NOMCORTO), Adocatprod.Recordset!NOMCORTO, "")
    
    'txtfechfin.Text = IIf(Not IsNull(DateValue(Trim(Adocatprod.Recordset!fechaactivo))), DateValue(Trim(Adocatprod.Recordset!fechaactivo)), "01/01/01")
    Txtclavdprov.Text = IIf(Not IsNull(Adocatprod.Recordset!clavedelprov), Adocatprod.Recordset!clavedelprov, "")
    txtpeso.Text = IIf(Not IsNull(Adocatprod.Recordset!peso), Adocatprod.Recordset!peso, "")
    txtlargo.Text = IIf(Not IsNull(Adocatprod.Recordset!largo), Adocatprod.Recordset!largo, "")
    txtancho.Text = IIf(Not IsNull(Adocatprod.Recordset!ancho), Adocatprod.Recordset!ancho, "")
    txtalto.Text = IIf(Not IsNull(Adocatprod.Recordset!Alto), Adocatprod.Recordset!Alto, "")
    
    'Me.txtfamilant.Text = IIf(Not IsNull(Adocatprod.Recordset!clafamil), Adocatprod.Recordset!clafamil, "")
    Chknal.Value = Adocatprod.Recordset!procedencia
    chkProm.Value = IIf(Adocatprod.Recordset!enpromo, 1, 0)
    If Trim(Txtprecosca.Text) <> "" And Trim(Txtpzaxca.Text) <> "" Then
        If Val(Txtprecosca.Text) <> 0 And Val(Txtpzaxca.Text) <> 0 Then
        Txtcoprecio.Text = Round(Val(Txtprecosca.Text) / Val(Txtpzaxca.Text), 2)
        
        Else
        Txtcoprecio.Text = ""
        End If
    End If
    'MsgBox Adocatprod.Recordset!tipoproducto
    'moy
    cmbtipo.Text = IIf(Not IsNull(Adocatprod.Recordset!tipoproducto), Adocatprod.Recordset!tipoproducto, "  ")
    
    'poner la ultima fecha de actualizacion
    fechaact.Caption = IIf(Not IsNull(Adocatprod.Recordset!fecact), Adocatprod.Recordset!fecact, "99/99/99")
    'USUARIO
    lbusuario.Caption = "MODIFICO : " & IIf(Not IsNull(Adocatprod.Recordset!USUARIO), Adocatprod.Recordset!USUARIO, "    ")
    ''lleno combo de lineas de este proveedor
    'Adofamilia.CommandType = adCmdText
    'Adofamilia.CursorType = adOpenKeyset
    'Adofamilia.ConnectionString = strconnect
    'cad = "select * from familias,lineas,departamento,linprove where fclave = clfamilia and sffamilia = fclave AND fdepto = depclave AND clprove = '" & Trim(Txtclaprove.Text) & "'"
    'Adofamilia.RecordSource = cad
    'Adofamilia.Refresh
    'Call llenafamilia
                  
    Chkcosituacion.Value = IIf(Adocatprod.Recordset!activo = True, 1, 0)
    RSTEMP.Open "SELECT * FROM CATPROV WHERE prove = '" & Trim(Txtclaprove.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If RSTEMP.EOF = False Then
        Cmbprove.Text = RSTEMP!NOMPROVE & "  Clave [" & RSTEMP!prove & "]"
    Else
        Cmbprove.Text = ""
    End If
    RSTEMP.Close
    RSTEMP.Open "SELECT * FROM lineas WHERE  sfclave = '" & Txtclafamil.Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If RSTEMP.EOF = False Then
        'SE TIENE QUE BUSCAR LA LINEA
        CAD = "sfclave = '" & Txtclafamil.Text & "'"
        If Adofamilia.Recordset.EOF = False And Adofamilia.Recordset.BOF = False Then
            Adofamilia.Recordset.MoveFirst
            Adofamilia.Recordset.Find CAD
        Else
            MsgBox "Esta familia no está relacionada con el proveedor"
        End If
        If Adofamilia.Recordset.EOF = False Then
            Cmbfamilia.Text = Adofamilia.Recordset!sfdescrip & "  Clave [" & Adofamilia.Recordset!sfclave & "]"
            Lbfam.Caption = Adofamilia.Recordset!fdescrip
            Lbldepto.Caption = Adofamilia.Recordset!depdescrip
        Else
            MsgBox "Esta familia no está relacionada con el proveedor"
        End If
    Else
        Cmbfamilia.Text = ""
        Lbfam.Caption = ""
        Lbldepto.Caption = ""
    End If
    RSTEMP.Close
    
    Adoprecio.CommandType = adCmdText
    Adoprecio.LockType = adLockOptimistic
    Adoprecio.CursorType = adOpenKeyset
    Adoprecio.RecordSource = "SELECT *  FROM preprod WHERE preclave = '" & Trim(Txtconsec.Text) & "' "
    Adoprecio.ConnectionString = strconnect
    Adoprecio.Refresh
    For i = 0 To 3
     Txtprecio(i).Text = ""
    Next
    
    If Adoprecio.Recordset.EOF = False And Adoprecio.Recordset.BOF = False Then
      '  Txtprecio(0).Text = Adoprecio.Recordset!precio1
      '  Txtprecio(1).Text = Adoprecio.Recordset!precio2
      '  Txtprecio(2).Text = Adoprecio.Recordset!precio3
      '  Txtprecio(3).Text = Adoprecio.Recordset!precio4
    End If
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtfechint_GotFocus()
'On Error GoTo error:
 '       If Txtfechint.Text = "" Then
 '         Cal1(0).Value = Trim(Str(Month(Date))) + "/" + Trim(Str(Day(Date))) + "/" + Mid(Trim(Str(Year(Date))), 3, 2)
 '         Cal1(0).Refresh
 '       Else
 '         Cal1(0).Value = Txtfechint.Text
 '       End If
 '       Cal1(0).Visible = True
 '       Cal1(0).SetFocus
 '      Exit Sub
'error:
'MsgBox Err.Description
End Sub
Private Sub dhabilitar()
'On Error GoTo error:
 Cmdbof.Enabled = False
 cmdcancela.Enabled = True
 cmdnuevo.Enabled = False
 cmdmodifica.Enabled = False
 Cmdnext.Enabled = False
 CmdBack.Enabled = False
 Cmdeof.Enabled = False
 Cmdbus.Enabled = False
 Cmdact.Enabled = True
 
Exit Sub
Error:
MsgBox Err.Description
End Sub
Private Sub habilitar()
'On Error GoTo error:
 Cmdbof.Enabled = True
 cmdcancela.Enabled = False
' cmdnuevo.Enabled = True
 cmdmodifica.Enabled = True
 Cmdnext.Enabled = True
 CmdBack.Enabled = True
 Cmdeof.Enabled = True
 Cmdact.Enabled = False
 Cmdbus.Enabled = True
 
Exit Sub
Error:
MsgBox Err.Description
 End Sub

Private Sub nuevoprod()
Dim i As Integer
On Error GoTo Error:

Dim rs As ADODB.Recordset
    
    'Txtconsec = Trim(Str(Val(rs.Fields!consec) + 1))
    Txtconsec.Text = ""
    Txtbarraspza.Text = ""
    Txtbarrascaja.Text = ""
    Txtdescripc.Text = ""
    Txtcontenido.Text = ""
'    Cmbmedida.Text = ""
    Cmbfamilia.Text = ""
    cmbtipo.Text = ""
    Cmbprove.Text = ""
    Lbldepto.Caption = ""
    Txtcoprecio.Text = ""
    Chkcosituacion.Value = 0
    'txtfechint.Text = Date
    'txtfechfin.Text = ""
    Txtclaprove.Text = ""
    Txtclafamil.Text = ""
    Txtnomcor.Text = ""
    Txtprecosca.Text = ""
    Txtpzaxca.Text = ""
    txtpeso.Text = "0"
    txtlargo.Text = "0"
    txtancho.Text = "0"
    txtalto.Text = "0"
    'txtfamilant.Text = " "
    Txtcoprecio.Text = ""
    For i = 0 To 3
        Txtprecio(i).Text = ""
    Next
    Txtprecosca.Enabled = True
    Txtcoprecio.Enabled = True
Exit Sub
Error:
MsgBox Err.Description
    
End Sub

Private Sub generamaximo()
Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    'rs.Source = "select max(consec) as consec from tfproduc where len(consec) > 6"
    rs.Source = "select max(consec) as consec from tfproduc WHERE consec >= 4600000"
    rs.ActiveConnection = cn
    rs.Open
   'MsgBox "El Producto ha sido dado de alta con la siguiente clave " & rs!CONSEC + 1, vbInformation
    If IsNull(rs!CONSEC) Then
       Txtconsec.Text = "4600000"
    Else
       Txtconsec.Text = rs!CONSEC + 1
    End If
   rs.Close
End Sub

Private Sub Txtpzaxca_KeyPress(KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub

Private Sub Txtpzaxca_LostFocus()
'On Error GoTo error:
If Trim(Txtpzaxca.Text) <> "" And Trim(Txtprecosca.Text) <> "" Then
Txtcoprecio.Text = Round(Val(Txtprecosca.Text) / Val(Txtpzaxca.Text), 2)
Else
Txtcoprecio.Text = Txtprecosca.Text
End If

Exit Sub
Error:
MsgBox Err.Description
End Sub
Private Sub llenafamilia()
Cmbfamilia.Clear
If Adofamilia.Recordset.EOF = False Then
Do While Not Adofamilia.Recordset.EOF

               If Not IsNull(Adofamilia.Recordset!sfclave) Then
               Cmbfamilia.AddItem Adofamilia.Recordset!sfdescrip + "  Clave [" + Adofamilia.Recordset!sfclave + "]"
               End If
              Adofamilia.Recordset.MoveNext
            Loop
           Adofamilia.Recordset.MoveFirst
End If
End Sub


