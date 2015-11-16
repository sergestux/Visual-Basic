VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmprecios 
   Caption         =   "ACTUALIZACION DE PRECIOS"
   ClientHeight    =   8595
   ClientLeft      =   480
   ClientTop       =   795
   ClientWidth     =   11370
   Icon            =   "frmprecios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin VB.Frame Fracontra 
      BackColor       =   &H00C0C000&
      Caption         =   "Contraseña de acceso"
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
      Height          =   1695
      Left            =   3240
      TabIndex        =   75
      Top             =   5160
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdConCanc 
         Caption         =   "&Cancelar"
         Height          =   350
         Left            =   2160
         TabIndex        =   80
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdConAce 
         Caption         =   "Aceptar"
         Height          =   350
         Left            =   600
         TabIndex        =   79
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtcontra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   76
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdcabecera 
      Caption         =   "Cab&ecera"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6480
      MaskColor       =   &H0080FFFF&
      TabIndex        =   134
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin MSAdodcLib.Adodc adodcprecio 
      Height          =   330
      Left            =   4680
      Top             =   8160
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
      Caption         =   "precios"
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
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   81
      Top             =   5160
      Width           =   11655
      Begin VB.TextBox Txtes 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1560
         TabIndex        =   93
         Top             =   1200
         Width           =   1100
      End
      Begin VB.TextBox Txtes 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   2640
         TabIndex        =   94
         Top             =   1200
         Width           =   1100
      End
      Begin VB.TextBox Txtes 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   3720
         TabIndex        =   95
         Top             =   1200
         Width           =   1100
      End
      Begin VB.TextBox Txtes 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   4800
         TabIndex        =   96
         Top             =   1200
         Width           =   1100
      End
      Begin VB.TextBox Txtes 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   5880
         TabIndex        =   97
         Top             =   1200
         Width           =   1100
      End
      Begin VB.TextBox Txtes 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   6960
         TabIndex        =   98
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CheckBox chkmanual 
         Caption         =   "Precios Manualmente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   8760
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         ToolTipText     =   "Observaciones decto. financiero 1"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   8160
         TabIndex        =   82
         ToolTipText     =   "Descuento financiero 1"
         Top             =   960
         Width           =   585
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   8160
         TabIndex        =   84
         ToolTipText     =   "Descuento financiero 2"
         Top             =   1320
         Width           =   585
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   8760
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         ToolTipText     =   "Observaciones decto. financiero 2"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   8160
         TabIndex        =   86
         ToolTipText     =   "Descuento financiero 3"
         Top             =   1680
         Width           =   585
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   8760
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         ToolTipText     =   "Observaciones decto. financiero 3"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   8160
         TabIndex        =   88
         ToolTipText     =   "Descuento financiero 4"
         Top             =   2040
         Width           =   585
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   8760
         ScrollBars      =   2  'Vertical
         TabIndex        =   89
         ToolTipText     =   "Observaciones decto. financiero 4"
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtDectoFin 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   8160
         TabIndex        =   90
         ToolTipText     =   "Descuento financiero 5"
         Top             =   2400
         Width           =   585
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   8760
         ScrollBars      =   2  'Vertical
         TabIndex        =   91
         ToolTipText     =   "Observaciones decto. financiero 5"
         Top             =   2400
         Width           =   2775
      End
      Begin MSMask.MaskEdBox Mskprecio 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   99
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecio 
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   102
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecio 
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   103
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecio 
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   104
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecioa 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   111
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecioa 
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   112
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecioa 
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   113
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecioa 
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   114
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecioa 
         Height          =   375
         Index           =   6
         Left            =   6120
         TabIndex        =   115
         Top             =   -120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecio 
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   100
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecioa 
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   127
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskpreciopza 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   105
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskpreciopza 
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   108
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskpreciopza 
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   109
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskpreciopza 
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   110
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskpreciopza 
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   106
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecio 
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   101
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskprecioa 
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   136
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskpreciopza 
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   107
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1/2  Caja Bodega"
         Height          =   495
         Index           =   5
         Left            =   3735
         TabIndex        =   137
         ToolTipText     =   "Venta de medio mayoreo autoservicio"
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio Pieza $"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   132
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1/2  Caja Envío"
         Height          =   495
         Index           =   4
         Left            =   2655
         TabIndex        =   128
         ToolTipText     =   "Venta de medio mayoreo autoservicio"
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio $"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   126
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pieza Autoservicio"
         Height          =   495
         Index           =   0
         Left            =   1560
         TabIndex        =   125
         ToolTipText     =   "Venta en Autoservicio"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mayoreo Intermedio"
         Height          =   495
         Index           =   2
         Left            =   5880
         TabIndex        =   123
         ToolTipText     =   "Venta mayoreo a domicilio"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mayoreo en Bodega"
         Height          =   495
         Index           =   3
         Left            =   6960
         TabIndex        =   122
         ToolTipText     =   "Venta mayoreo en bodega"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   120
         Picture         =   "frmprecios.frx":030A
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Escala %"
         Height          =   375
         Left            =   120
         TabIndex        =   121
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Lblfechact 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   8160
         TabIndex        =   120
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio  Ant. $"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   119
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label txtusuario 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3600
         TabIndex        =   118
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   8880
         TabIndex        =   117
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "Decto.Fin."
         Height          =   255
         Left            =   8160
         TabIndex        =   116
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "May. Envío y/o crédito"
         Height          =   495
         Index           =   1
         Left            =   4800
         TabIndex        =   124
         ToolTipText     =   "Venta mayoreo autoservicio"
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Fraautoriza 
      BackColor       =   &H00808000&
      Caption         =   "Clave de Autorizacion..."
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4800
      TabIndex        =   77
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtautoriza 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   78
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton mediomay 
      Caption         =   "&M.M."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7560
      MaskColor       =   &H0080FFFF&
      TabIndex        =   62
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10800
      TabIndex        =   67
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9720
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   66
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbprod 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   60
      Top             =   -120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmdgrabar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8640
      MaskColor       =   &H0080FFFF&
      TabIndex        =   64
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtclprod 
      Height          =   315
      Left            =   3000
      TabIndex        =   19
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11655
      Begin VB.Label lblproducto 
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   600
         Width           =   5415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Familia :"
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
         Left            =   6720
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Depto   :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lblpres 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   960
         Width           =   5055
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label21 
         Caption         =   "Present.    :"
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
         TabIndex        =   35
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Lbllinea 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7680
         TabIndex        =   34
         Top             =   960
         Width           =   3855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Linea    :"
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
         Left            =   6720
         TabIndex        =   31
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Producto   :"
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
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblprov 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   1320
         TabIndex        =   23
         Top             =   240
         Width           =   5415
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblfamilia 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   7680
         TabIndex        =   22
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lbldepto 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   7680
         TabIndex        =   20
         Top             =   240
         Width           =   3615
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   11655
      Begin VB.TextBox TxtDepto 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   10920
         TabIndex        =   138
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox Mskpreprom 
         Height          =   495
         Left            =   7560
         TabIndex        =   129
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcprod 
         Height          =   330
         Left            =   6600
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
      Begin VB.TextBox txtpreciodes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   -120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Txtpreciocargo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   -120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Promociones :"
         Height          =   735
         Left            =   8160
         TabIndex        =   26
         Top             =   1320
         Width           =   3375
         Begin VB.TextBox txtppca 
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
            Left            =   2400
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtpnoca 
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
            Left            =   840
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "No. cajas "
            Height          =   345
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "En Cada"
            Height          =   345
            Left            =   1680
            TabIndex        =   37
            Top             =   240
            Width           =   735
         End
      End
      Begin MSAdodcLib.Adodc Adodccargo 
         Height          =   330
         Left            =   1920
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   3
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
         DataSourceName  =   "pitico"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "cargos"
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
      Begin MSAdodcLib.Adodc adodcdes 
         Height          =   330
         Left            =   1560
         Top             =   2760
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
         Caption         =   "descuentos"
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
      Begin VB.Frame Frame7 
         Caption         =   "Cargos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Width           =   3855
         Begin VB.TextBox Txtcar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   8
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox Txtcar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   7
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox Txtcar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   6
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox Txtcar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   5
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Txtcar 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   4
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Txtcar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2520
            MaxLength       =   5
            TabIndex        =   3
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox Txtcar 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2520
            MaxLength       =   5
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 13    Cargo por  Maniobras  $"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 12       Flete efectivo           $"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 5        Otro cargo efectivo   $"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 4                     IVA                %"
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 3                    IEPS                %"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 1                   Cargo1              %"
            Height          =   375
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 2                   Cargo 2             %"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Descuentos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   4200
         TabIndex        =   27
         Top             =   120
         Width           =   3855
         Begin VB.TextBox Txtdes 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   2640
            MaxLength       =   5
            TabIndex        =   13
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox Txtdes 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   15
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox Txtdes 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   14
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox Txtdes 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2640
            MaxLength       =   5
            TabIndex        =   12
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Txtdes 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2640
            MaxLength       =   5
            TabIndex        =   11
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Txtdes 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2640
            MaxLength       =   5
            TabIndex        =   10
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox Txtdes 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2640
            MaxLength       =   5
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 10           Descuento 5           %"
            Height          =   375
            Left            =   240
            TabIndex        =   54
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 12     Descuento Financiero"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 11     Descuento Efectivo       $"
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 9             Descuento 4           %"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 8             Descuento 3           %"
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 7             Descuento 2           %"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 6             Descuento 1           %"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Width           =   2415
         End
      End
      Begin MSMask.MaskEdBox mskcostocaja 
         Height          =   375
         Left            =   10080
         TabIndex        =   56
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskcostopza 
         Height          =   375
         Left            =   10080
         TabIndex        =   57
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lbltasa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEPARTAMENTO CON TASA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8400
         TabIndex        =   135
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label OFERTADO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OFERTADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   9240
         TabIndex        =   133
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo Pza:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8520
         TabIndex        =   58
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo Caja:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8520
         TabIndex        =   53
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Lbldescripd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   7560
         TabIndex        =   25
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "costo con prom."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   130
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adoescala 
      Height          =   330
      Left            =   3480
      Top             =   -120
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
      Caption         =   "escalas"
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
   Begin VB.Frame frmmm 
      Caption         =   "Medio Mayoreo..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   600
      TabIndex        =   63
      Top             =   720
      Visible         =   0   'False
      Width           =   10815
      Begin VB.Frame frmalta 
         Caption         =   "Alta"
         Enabled         =   0   'False
         Height          =   2175
         Left            =   2640
         TabIndex        =   71
         Top             =   1320
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtpaquetesmm 
            Height          =   495
            Left            =   960
            TabIndex        =   72
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label10 
            Caption         =   "Paquetes por Caja"
            Height          =   255
            Left            =   1080
            TabIndex        =   73
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.CommandButton cmdmm 
         Caption         =   "&Regresar"
         Height          =   495
         Index           =   2
         Left            =   6360
         TabIndex        =   70
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdmm 
         Caption         =   "&Eliminar"
         Height          =   495
         Index           =   1
         Left            =   4800
         TabIndex        =   69
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdmm 
         Caption         =   "&Nuevo"
         Height          =   495
         Index           =   0
         Left            =   3240
         TabIndex        =   68
         Top             =   4560
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid gridmm 
         Bindings        =   "frmprecios.frx":0614
         Height          =   3735
         Left            =   360
         TabIndex        =   65
         Top             =   600
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "consec"
            Caption         =   "Clave"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#.##0 ""Pta"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descripc"
            Caption         =   "Descripcion"
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
            DataField       =   "Paquetes"
            Caption         =   "Paquetes"
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
            DataField       =   "Precio1"
            Caption         =   "Precio"
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
         BeginProperty Column04 
            DataField       =   "activo"
            Caption         =   "Activo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   4995.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adomm 
         Height          =   495
         Left            =   240
         Top             =   4560
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
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
         Caption         =   "Adomm"
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
   Begin MSMask.MaskEdBox txtprecosto 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Precio Lista "
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
      Left            =   240
      TabIndex        =   131
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label30 
      Caption         =   "Label30"
      Height          =   615
      Left            =   480
      TabIndex        =   74
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label prov 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "moy"
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
      Left            =   240
      TabIndex        =   61
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Lblcodbarra 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   59
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmprecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim costodescprod As Double
Dim PRECOSTO As Double
Dim costotal As Double
Private npzaxcaj As Integer
Private fecha1 As Date
Private fecha2 As Date
Private lsibusca As Boolean
Private lfecha As Boolean
Private antprecio As Double
Private altamm As Boolean
Private valorman As String
Private nValAnt As Double

Private Sub cmbprod_Click()
On Error GoTo Error:
If lsibusca Then
If Trim(cmbprod.Text) <> "" Then
    If Trim(Lbldepto.Caption) = "" Then
        Dim Valor As Integer
        Dim strclave As String
        Dim valor2 As Integer
        Valor = InStr(1, cmbprod.Text, "\")
        strclave = Mid(cmbprod.Text, Valor + 1, Len(cmbprod.Text) - Valor - 1)
        adodcprod.Recordset.Bookmark = Val(strclave)
        cmbprod.Visible = False
        cmdnuevo.Visible = True
        cmdsalir.Visible = True
        cmdGrabar.Visible = True
        txtclprod.Locked = True
        Call asigna
    End If
Else
    cmbprod.SetFocus
End If
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Function VALIDAPRECIOS() As Boolean
'MsgBox precosto
VALIDAPRECIOS = True
If mskcostocaja = 0 Then
  Call calculo
End If
costotempo = mskcostocaja.Text
'CONDICIONES EL COSTO DEBE SER MAYOR A TODOS LOS PRECIOS
If Val(strcostotempo) > Val(Mskprecio(1).Text) Then
   MsgBox "El precio de Venta  es Menor  al costo "
   VALIDAPRECIOS = False
End If
If Val(costotempo) > Val(Mskprecio(3).Text) Then
   MsgBox "El precio de Venta  es Menor  al costo ", vbInformation
   VALIDAPRECIOS = False
End If

If tipotienda = 1 Then
    If Val(Mskprecio(0).Text) < Val(Mskprecio(4).Text) Then
        MsgBox "EL PRECIO DE MEDIO MAYOREO DEBE SER MENOR AL PRECIO DE AUTOSERVICIO", vbExclamation, "SIPITICO"
        VALIDAPRECIOS = False
    End If
    If Val(Mskprecio(1).Text) / adodcprod.Recordset!PAQUETES > Val(Mskprecio(4).Text) Then
        MsgBox "EL PRECIO DE MEDIO MAYOREO DEBE SER MAYOR AL PRECIO DE MAYOREO ENTRE PAQUETES", vbExclamation, "SIPITICO"
        VALIDAPRECIOS = False
    End If
    If Val(mskcostopza.Text) > Val(Mskprecio(2).Text) Then
        MsgBox "El precio de Venta de Medio Mayoreo es Menor  al costo "
        VALIDAPRECIOS = False
    End If
End If

'3a CONDICION EL PRECIO DE ESCALA 2 DEBE SER MAYOR A LA ESCALA 4
'If Val(Mskprecio(2).Text) <= Val(Mskprecio(3).Text) Then
'   MsgBox "El precio de Escala 3 debe ser mayor a precio de Escala 4"
'   VALIDAPRECIOS = False
'End If

'VALIDACION DE QUE DEBE TENER EL IVA DE ACUERDO A LA ZONA
'If Val(Txtcar(3).Text) > 0 And Val(Txtcar(3).Text) < 15 Then
'MsgBox Txtcar(3).Text
'If Val(Txtcar(3).Text) = 0 Then
'    VALIDAPRECIOS = True
'ElseIf Val(Txtcar(3).Text) = 15 Then
'   VALIDAPRECIOS = True
'Else
'   VALIDAPRECIOS = False
'End If
End Function

Private Sub cmdcabecera_Click()
If MsgBox("CONFIRMA SI DESEAS GENERAR ARCHIVO DE MICROSOFT WORD CON LETERERO PARA CABECERA", vbQuestion + vbYesNo) = vbYes Then
    Dim ApDoc As Word.Application
    Dim rs As ADODB.Recordset
    Dim N As Integer
    Dim dec As String
    On Error GoTo Error:
    Cmdlg.DialogTitle = "Ruta donde se grabará el archivo de cabeceras"
    Cmdlg.InitDir = "C:\"
    Cmdlg.Filter = "Archivos Microsoft Word (*.doc) | *.doc"
    Cmdlg.CancelError = True
    Cmdlg.ShowSave
    Set ApDoc = CreateObject("word.Application")  'run it
    ApDoc.Visible = True
    ApDoc.Documents.Open FileName:=App.Path & "\cabecera.doc", ConfirmConversions:=False, _
        ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto
'ApDoc.ActiveDocument.Shapes("Picture 5").Select
'ApDoc.Selection.Copy
'ApDoc.Documents.Close False
'With ApDoc
'    .Documents.Add DocumentType:=wdNewBlankDocument
'    .ActiveDocument.PageSetup.Orientation = wdOrientPortrait
'    .ActiveDocument.PageSetup.PaperSize = wdPaperLegal
'End With

     With ApDoc
        '.Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        '.Selection.EndKey Unit:=wdLine  'Agrego linea
        '.Selection.TypeParagraph
        .Selection.Font.Name = "Benguiat Bk BT"
        .Selection.Font.Size = 36
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:=lblProducto.Caption
        .Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        .Selection.EndKey Unit:=wdLine  'Agrego linea
        .Selection.TypeParagraph
        .Selection.Font.Name = "Compacta Bd BT"
        .Selection.Font.Size = 270
        .Selection.Font.Shadow = True
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         nPos = InStr(1, Mskprecio(0).Text, ".")
         If nPos > 0 Then
           .Selection.TypeText Text:=Mid(Mskprecio(0).Text, 1, nPos - 1) & " "
         Else
           .Selection.TypeText Text:=Mskprecio(0).Text & " "
         End If
        .Selection.Font.Superscript = True
        .Selection.Font.Underline = wdUnderlineSingle
        .Selection.Font.Size = 245
        'MsgBox Mskprecio(0).Text
        If (nPos > 0) Then
           MsgBox Mid(Mskprecio(0).Text, nPos)
           dec = Mid(Mskprecio(0).Text, nPos)
           .Selection.TypeText Text = dec
        Else
           .Selection.TypeText Text:="00"
        End If
        .Selection.Font.Name = "BenguiatGot Bk BT"
        .Selection.Font.Size = 26
        .Selection.EndKey Unit:=wdLine  'Agrego linea
        .Selection.TypeParagraph
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="- " & lblcodbarra.Caption & " -"
        Exit Sub

        .Selection.Font.Shadow = False
        .Selection.TypeParagraph
        .Selection.Font.Size = 85
        .Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        .Selection.EndKey Unit:=wdLine  'Agrego linea
        .Selection.TypeText Text:=Format(rs!precio1, "$ ###,###,#0.00")
        .Selection.EndKey Unit:=wdLine  'Agrego linea
        .Selection.Font.Size = 18
        .Selection.Font.Outline = True
        .Selection.TypeParagraph
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="VIGENCIA DEL " & UCase(Format(DateAdd("d", 1, rs!FechaIni), "dd MMM.")) & " AL " & UCase(Format(rs!FechaFin, "dd MMM."))
        .Selection.TypeParagraph
        .Selection.TypeText Text:="O HASTA AGOTAR EXISTENCIAS"
        .Selection.EndKey Unit:=wdStory  'Fin de seccion 2
        If (N + 1) Mod 2 = 0 And N > 0 Then
            .Selection.InsertBreak Type:=wdPageBreak
        Else
            .Selection.Paste
            .Selection.EndKey Unit:=wdStory
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
        End If
        N = N + 1
    End With
    rs.MoveNext
ActiveDocument.SaveAs FileName:=Cmdlg.FileName, FileFormat:=wdFormatDocument _
        , LockComments:=False, Password:="", AddToRecentFiles:=True, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
         SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False
ApDoc.Quit
Set ApDoc = Nothing
End If
Exit Sub
Error:
   MsgBox Err.Description, vbCritical
End Sub


Private Sub cmdConAce_Click()
On Error GoTo Error:
Dim RsCon As ADODB.Recordset
'SE VALIDAN LOS DATOS
' If txtContra.Text = "MODI12" Then
 Set RsCon = New ADODB.Recordset
 If Trim(txtContra.Text) = "" Then Exit Sub
 RsCon.Open "SELECT * FROM USUARIOS WHERE PASSMASTER = '" & Trim(txtContra.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
 If RsCon.RecordCount = 0 Then
    MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
    txtContra.SetFocus
    SendKeys "+{HOME}"
    'chkmanual.Value = 0
    Exit Sub
Else
 For i = 0 To 5
    Txtes(i).Enabled = True
    Mskprecio(i).Enabled = True
 Next
End If
 FRACONTRA.Visible = False
 Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdConCanc_Click()
    FRACONTRA.Visible = False
    chkmanual.Value = 0
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo Error:
Dim rsttemp As ADODB.Recordset
Dim Strfecha1 As String
Dim Strfecha2 As String

'ANTES QUE NADA SE DEBE VALIDAR QUE LOS PRECIOS SEAN MAYORES AL COSTO
'SE VALIDA PRIMERO QUE NO HAYA DIFERENCIAS ENTRE EL COSTO QUE SE TRAJO
'CON EL COSTO QUE SE PONE EN LA MASCARILLA
'If precosto <> mskcostocaja Then
'   MsgBox "Upps se debe mover el mouse"
'End If
If tipotienda <> 2 And tipotienda <> 4 Then
   If Not puedegrabar Then
      If Val(Txtes(0).Text) < 12 Then
        MsgBox "PARA REALIZAR ESTE CAMBIO DEBE PROPORCIONAR LA CLAVE    ", vbInformation
        Me.Fraautoriza.Enabled = True
        Me.Fraautoriza.Visible = True
        'Me.txtautoriza.SetFocus
        Exit Sub
      End If
    End If
End If

If Not VALIDAPRECIOS Then
   MsgBox "VERIFIQUE LOS PRECIOS, NO ESTAN CORRECTOS !!", vbCritical
   Exit Sub
End If
If adodcprod.Recordset!OFERTADO = 0 Then
   respsn = MsgBox("REALMENTE DESEA GUARDAR LOS CAMBIOS ?", vbExclamation + vbYesNo, "Precios")
     
   If respsn = vbYes Then
        If PRECOSTO = 0 Then
           'SE NECESITA HACER UN RECALCULO
           Call calculo
        End If
       cn.Execute "update cargos set cargo1 = " & Txtcar(0).Text & ", cargo2 = " & Txtcar(1).Text & ",  " & _
                "ieps = " & Txtcar(2).Text & ",iva = " & Txtcar(3).Text & ", cargo_efectivo = " & Txtcar(4).Text & ", flete_efectivo = " & Txtcar(5).Text & ", maniobras = " & Txtcar(6).Text & _
                " WHERE caprod = '" & Trim(txtclprod.Text) & "'"
       CADENA = "update descuentos set deprod = '" & Trim(txtclprod.Text) & "' , decto1 = " & Txtdes(0).Text & " , decto2 = " & Txtdes(1).Text & " ,  " & _
               " decto3 = " & Txtdes(2).Text & ",dectooferta = " & Txtdes(3).Text & " , dectofinanciero = " & Txtdes(4).Text & " , dectoefectivo = " & Txtdes(5).Text & _
                " ,decto5 = " & Txtdes(6).Text & " WHERE deprod = '" & Trim(txtclprod.Text) & "'"
                 
       cn.Execute CADENA
       CADENA = " UPDATE preprod SET precio1 = " & Mskprecio(0).Text & ", precio2 = " & Mskprecio(1).Text & ", precio3 = " & Mskprecio(2).Text & ", precio4 = " & Mskprecio(3).Text & ", precio5 = " & Mskprecio(4).Text & ", precio6 = " & Mskprecio(5).Text & _
                     ", fechaact = '" & date & "'" & ",prusuario =  '" & cUsuario & "'" & _
                    " WHERE preclave = '" & Trim(txtclprod.Text) & "'"
       'ACTUALIZACION DE LA FECHA Y USUARIO
       Me.Lblfechact.Caption = "Fecha Ult. Act.  " & date + Time
       Me.Lblfechact.Refresh
       'MsgBox CADENA
       cn.Execute CADENA
       adodcprecio.Refresh
       adodcprecio.Recordset!fechaact = date + Time
       adodcprecio.Recordset.Update

       'VALIDAR LA TASA IEPS, parametros : IEPS,IVA
       'tasatempo = validatasa(Txtcar(2).Text, Txtcar(3).Text)
       tasatempo = Trim(Txtdepto.Text)
       If Me.chkmanual = 1 Then
          valorman = "1"
       Else
          valorman = "0"
       End If
       pzasmay = IIf(adodcprod.Recordset!PAQUETES >= 10, Round(adodcprod.Recordset!PAQUETES / 3 + 0.1), 0)
       CADENA = "UPDATE tfproduc SET flesub = " & pzasmay & ", fletex = " & IIf(Mskprecio(4).Text = "", 0, Mskprecio(4).Text) & ", precosto = " & Str(Round(PRECOSTO, 2)) & ", costocaj = " & txtprecosto.Text & ", cajas = " & txtpnoca.Text & ", encajas = " & txtppca.Text & ", costopaq = " & Str(Round(PRECOSTO, 2)) / adodcprod.Recordset!PAQUETES & ",fecact =  " & _
                 "'" & date & "' , actualizado = 1 , tasaieps = " & tasatempo & ", manual = '" & valorman & "' , IVA = " & Txtcar(3).Text & ", IEPS =  " & Txtcar(2).Text & _
                 ", dectofin1 = '" & txtDectoFin(0).Text & "', observa1 = '" & txtobserva(0).Text & "', dectofin2 = '" & txtDectoFin(1).Text & "', observa2 = '" & txtobserva(1).Text & "', dectofin3 = '" & txtDectoFin(2).Text & "', observa3 = '" & txtobserva(2).Text & "', dectofin4 = '" & txtDectoFin(3).Text & "', observa4 = '" & txtobserva(3).Text & "', dectofin5 = '" & txtDectoFin(4).Text & "', observa5 = '" & txtobserva(4).Text & "' " & _
                 " WHERE consec = '" & Trim(txtclprod.Text) & "'"
       cn.Execute CADENA
       cn.Execute "UPDATE margen SET escala1 = " & Txtes(0).Text & " , escala2 = " & Txtes(1).Text & " , escala3 = " & Txtes(2).Text & " , escala4 = " & Txtes(3).Text & " , escala5 = " & Txtes(4).Text & " , escala6 = " & Txtes(5).Text & _
                       " WHERE producto = '" & Trim(txtclprod.Text) & "'"
       cn.Execute "UPDATE descprod SET  costo = " & costodescprod & ",PReciolista=" & Val(txtprecosto.Text) & _
                  " WHERE Producto = '" & Trim(txtclprod.Text) & "'"
        
       Set rsttemp = New ADODB.Recordset
       rsttemp.Open "SELECT * FROM Descprod WHERE Producto =  '" & Trim(txtclprod.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     
       If rsttemp.BOF And rsttemp.EOF Then
                cn.Execute "INSERT INTO descprod (proveedor,producto,decto1,decto2,decto3,decto4,financiero," & _
                "efectivo,cargo1,cargo2,cargo3,cargo4,cargo5,cajas,encajas,costo,preciolista,flete,maniobras,plazopago,situacion, fechaact)" & _
                " VALUES ('" & adodcprod.Recordset!claprove & "','" & Trim(txtclprod.Text) & _
                "'," & Val(Txtdes(0).Text) & "," & Val(Txtdes(1).Text) & _
                "," & Val(Txtdes(2).Text) & ", " & Val(Txtdes(3).Text) & _
                "," & Val(Txtdes(5).Text) & "," & Val(Txtdes(4).Text) & _
                "," & Val(Txtcar(0).Text) & "," & Val(Txtcar(1).Text) & "," & _
                Val(Txtcar(3).Text) & "," & Val(Txtcar(2).Text) & "," & _
                Val(Txtcar(4).Text) & "," & Val(txtpnoca.Text) & "," & _
                Val(txtppca.Text) & "," & costodescprod & "," & _
                Val(txtprecosto.Text) & "," & Val(Txtcar(5).Text) & "," & _
                Val(Txtcar(6).Text) & ",0,1" & ",'" & date + Time & "')"
       Else
                CADENA = "UPDATE descprod SET proveedor = '" & adodcprod.Recordset!claprove & "',producto =  '" & Trim(txtclprod.Text) & "',decto1 = " & Val(Txtdes(0).Text) & ",decto2 =" & Val(Txtdes(1).Text) & ",decto3 =" & Val(Txtdes(2).Text) & "," & _
                "decto4 =" & Val(Txtdes(3).Text) & ",financiero=" & Val(Txtdes(5).Text) & ",efectivo=" & Val(Txtdes(4).Text) & ",cargo1 =" & Val(Txtcar(0).Text) & ",cargo2=" & Val(Txtcar(1).Text) & ",cargo3=" & Val(Txtcar(3).Text) & ",cargo4=" & Val(Txtcar(2).Text) & "," & _
                "cargo5=" & Val(Txtcar(4).Text) & ", costo = " & costodescprod & ", cajas = " & Val(txtpnoca.Text) & ", encajas = " & Val(txtppca.Text) & ", preciolista=" & Val(txtprecosto.Text) & ",situacion = 1,fechaact = '" & date + Time & "', flete =" & Val(Txtcar(5).Text) & ",maniobras =" & Val(Txtcar(6).Text) & _
                " WHERE Producto = '" & Trim(txtclprod.Text) & "'"
                'MsgBox cadena
                cn.Execute CADENA
       End If
       rsttemp.Close
       rsttemp.Open "SELECT * FROM cambpre WHERE producto = '" & Trim(txtclprod.Text) & "' AND modificado = 0", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
       If Not (rsttemp.BOF And rsttemp.EOF) Then
          cn.Execute "UPDATE cambpre SET modificado = 1 WHERE producto = '" & Trim(txtclprod) & "'"
       End If
   End If
Else
    MsgBox " Este producto se encuentra OFERTADO no se puede modificar ", vbCritical
End If
'se deben actualizar sus medios mayoreos con una sentencia sql
'costomm = (Mskprecio(0).Text) * 0.9
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "select PAQUETES  from tfproduc where consec = '" & Trim(txtclprod.Text) & "'"
rs.ActiveConnection = cn
rs.Open
paq = rs.Fields!PAQUETES
rs.Close
escalamm = Val(Txtes(0).Text) * 0.9
'se debe sacar el costo por caja
PREPZA = PRECOSTO / paq
costomm = ((PREPZA * escalamm) / 100) + PREPZA
CADENA = "update tfproduc set costocaj =  round(" & costomm & " *  paquetes,1), precosto = round( " & costomm & " * paquetes,1), fecact = '" & date & "' where medmay = '" & Trim(Me.txtclprod.Text) & "'"
cn.Execute CADENA
'y tambien la parte precios preprod
pretemp = costomm
CADENA = "update preprod set precio1 = precosto , precio2 = precosto, precio3 = precosto, precio4 = precosto from tfproduc,preprod where medmay = '" & Trim(Me.txtclprod.Text) & "' and consec = preclave "
cn.Execute CADENA
txtusuario.Caption = "Modifico: " & cUsuario
txtusuario.Refresh
Exit Sub
Error:
MsgBox Err.Description
End Sub

Function validatasa(TIEPS As Integer, TIVA As Integer) As Integer
'If TIVA = 0 And TIEPS = 0 Then
'   validatasa = 1
'ElseIf TIVA = 15 And TIEPS = 0 Then
'   If MsgBox("Es facturado como producción 2001" & Chr(13) & Chr(13) & "NO  => Producto con departamento con tasa 2" & Chr(13) & "SI    => Producto con departamento con tasa 5", vbQuestion + vbYesNo + vbDefaultButton2, "Departamento") = vbYes Then
'      validatasa = 5
'   Else
'      validatasa = 2
'   End If
'ElseIf TIVA = 15 And TIEPS = 25 Then
'   validatasa = 3
'ElseIf TIVA = 15 And TIEPS = 30 Then
'   validatasa = 4
'ElseIf TIVA = 15 And TIEPS = 50 Then
'   validatasa = 6
'ElseIf TIVA = 15 And TIEPS = 5 Then
'   validatasa = 7
'ElseIf TIVA = 15 And TIEPS = 20 Then
'   validatasa = 8
'Else
'   MsgBox "ESTE PRODUCTO NO SE ENCUENTRA DENTRO DE ALGUN DEPARTAMENTO CON TASA REGISTRADA, FAVOR DE INFORMAR AL ADMINISTRADOR DEL SISTEMA", vbCritical, "No se encuentra tasa registrada"
'End If
'lbltasa.Caption = "DEPARTAMENTO CON TASA " & validatasa
End Function

Private Sub cmdmm_Click(Index As Integer)
Select Case Index
Case 0
   frmalta.Visible = True
   frmalta.Enabled = True
   altamm = True
   Me.txtpaquetesmm.SetFocus
Case 1
   clave = Adomm.Recordset!CONSEC
   CADENA = "update tfproduc set activo = 0 where consec = '" & Trim(clave) & "'"
   'MsgBox cadena
   cn.Execute CADENA
   Call mediomay_Click
Case 2
   frmmm.Enabled = False
   frmmm.Visible = False
   'Adomm.Recordset.Close
End Select
End Sub

Private Sub nuevomm()
'se agrega un producto de medio mayoreo
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "select max(consec) as consec from tfproduc where len(consec) > 6"
rs.ActiveConnection = cn
rs.Open
clave = Trim(Str(Val(rs.Fields!CONSEC) + 1))
rs.Close
'con sentencia sql
costomm = (Mskprecio(0).Text * 0.9) * Val(Me.txtpaquetesmm.Text)
'MsgBox Mskprecio(0).Text
'MsgBox costomm
'DATOS GENERALES DEL PRODUCTO
CADENA = "INSERT INTO TFPRODUC(CONSEC,CLAPROVE,LINEA,DESCRIPC,CONTENID,NOMCORTO,ACTIVO,FECACT,PAQUETES,COSTOCAJ,medida,medmay)   VALUES(" & _
"'" & clave & "','" & Trim(prov.Caption) & "'," & 757 & ",'" & Trim(lblProducto.Caption + "          M.M.") & "'," & adodcprod.Recordset!Contenid & ",'" & _
Mid(lblProducto.Caption, 1, 20) & "',1" & ",'" & date & "'," & Trim(txtpaquetesmm.Text) & "," & costomm & ",'" & adodcprod.Recordset!medida & "','" & Trim(Me.txtclprod.Text) & "')"
'MsgBox cadena
cn.Execute CADENA
'ESCALAS EN 0
CADENA = "INSERT INTO MARGEN(PRODUCTO)   VALUES(" & "'" & clave & "')"
'MsgBox cadena
cn.Execute CADENA
'PRECIOS
CADENA = "INSERT INTO PREPROD(PRECLAVE,PRECIO1,PRECIO2,PRECIO3,PRECIO4) VALUES(" & "'" & clave & "'," & costomm & "," & costomm & "," & costomm & "," & costomm & ")"
'MsgBox cadena
cn.Execute CADENA
'MsgBox cadena
'DESCPROD DESCUENTOS POR DEFAULT
CADENA = "INSERT INTO DESCPROD(PRODUCTO)   VALUES(" & "'" & clave & "')"
'MsgBox cadena
cn.Execute CADENA
'CARGOS, AQUI SI SE PONE EL IVA Y IEPS
CADENA = "INSERT INTO CARGOS(CAPROD,IVA,IEPS)   VALUES(" & "'" & clave & "'," & Txtcar(3).Text & "," & Txtcar(2).Text & ")"
'MsgBox cadena

cn.Execute CADENA
'DESCUENTOS
CADENA = "INSERT INTO DESCUENTOS(DEPROD)   VALUES(" & "'" & clave & "')"
'MsgBox cadena
cn.Execute CADENA
Call mediomay_Click
End Sub
Private Sub cmdnuevo_Click()
On Error GoTo Error:
 lblProducto.Caption = ""
 lblProducto.Refresh
 lblProv.Caption = ""
 lblProv.Refresh
 lblfamilia.Caption = ""
 lblfamilia.Refresh
 Lbldepto.Caption = ""
 Lbldepto.Refresh
 lblcodbarra.Caption = ""
 lblcodbarra.Refresh
 Lbllinea.Caption = ""
 Lbllinea.Refresh
 lblpres.Caption = ""
 lblpres.Refresh
Frame2.Visible = False
txtclprod.Locked = False
txtclprod.SetFocus
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

Private Sub chkmanual_Click()
If chkmanual.Value = 1 Then
    FRACONTRA.Enabled = True
    FRACONTRA.Visible = True
    txtContra.Visible = True
    txtContra.Enabled = True
    If txtContra.Visible Then txtContra.SetFocus
End If
'MsgBox chkmanual.Value
End Sub

Private Sub Form_Activate()
On Error GoTo Error:
If adodcprod.Recordset!MEDMAY < 1 Then
   If adodcprod.Recordset!OFERTADO = False Then 'And adodcprod.Recordset!medmay < 1 Then
      txtprecosto.SetFocus
  End If
End If
If txtContra.Visible Then txtContra.SetFocus
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_Deactivate()
lpprov = False
'If Not VALIDAPRECIOS Then
'    MsgBox "Existen  Errores"
'    Cancel = True
'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 119
   frmCalc.Show 1
Case 27
    Unload Me
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
   If altamm = True Then
      Me.frmalta.Enabled = False
      Me.frmalta.Visible = False
       'procesa alta
      Call nuevomm
   End If
   KeyAscii = 0
   'SendKeys "{TAB}"
   keybd_event &H9, 0, 0, 0
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Error:
If Sql Then
   strcont = "SELECT tfproduc.usuario as usuario,medmay,manual,consec,claprove,cajas,encajas,precosto,barraspza,contenid,medida,descripc,costocaj,costopaq,paquetes,nomprove,sfdescrip,fdescrip,depdescrip,ofertado,fecact " & _
             ",tfproduc.dectofin1 as dectofin1, TFPRODUC.OBSERVA1 as observa1,TFPRODUC.DECTOFIN2 as dectofin2,TFPRODUC.OBSERVA2 as observa2,TFPRODUC.DECTOFIN3 as dectofin3,TFPRODUC.OBSERVA3 as observa3,TFPRODUC.DECTOFIN4 as dectofin4,TFPRODUC.OBSERVA4 as observa4,TFPRODUC.DECTOFIN5 as dectofin5,TFPRODUC.OBSERVA5 as observa5, tfproduc.tasaieps " & _
             "FROM tfproduc,catprov,familias,lineas,departamento " & _
             "WHERE claprove = prove AND LINEA = Sfclave AND sffamilia = fclave AND fdepto = depclave  and consec = '" & Trim(strcveprod) & "'  ORDER BY DESCRIPC"
Else
   strcont = "SELECT TFPRODUC.TASAIEPS,tfproduc.usuario as usuario,medmay,manual,consec,claprove,cajas,encajas,precosto,barraspza,contenid,medida,descripc,costocaj,costopaq,paquetes,nomprove,ofertado, fecact, '' AS sfdescrip, '' AS fdescrip,'' AS depdescrip" & _
             ",tfproduc.dectofin1 as dectofin1, TFPRODUC.OBSERVA1 as observa1,TFPRODUC.DECTOFIN2 as dectofin2,TFPRODUC.OBSERVA2 as observa2,TFPRODUC.DECTOFIN3 as dectofin3,TFPRODUC.OBSERVA3 as observa3,TFPRODUC.DECTOFIN4 as dectofin4,TFPRODUC.OBSERVA4 as observa4,TFPRODUC.DECTOFIN5 as dectofin5,TFPRODUC.OBSERVA5 as observa5 " & _
             "FROM tfproduc,catprov " & _
             "WHERE claprove = prove AND consec = '" & Trim(strcveprod) & "'  ORDER BY DESCRIPC"
End If
puedegrabar = False
lfecha = False
Dim adorsTemp As ADODB.Recordset

    adodcprod.CursorType = adOpenDynamic
    adodcprod.ConnectionString = cCadConex
    adodcprod.RecordSource = strcont
    adodcprod.Refresh
    cmbprod.Clear
    Do While Not adodcprod.Recordset.EOF
            If Not IsNull(adodcprod.Recordset!DESCRIPC) Then
                 des = adodcprod.Recordset!DESCRIPC
                 If IsNull(adodcprod.Recordset!PAQUETES) Then
                    paq = 0
                 Else
                    paq = Str(adodcprod.Recordset!PAQUETES)
                End If
                 pre = " (Presentacion : " + Str(paq) + "X  " + Str(adodcprod.Recordset!Contenid) + " "
                 med = Trim(adodcprod.Recordset!medida)
                 cla = "Clave [" + adodcprod.Recordset!CONSEC + " ] + \ " + Trim(Str(adodcprod.Recordset.Bookmark)) + ")"
                 cmbprod.AddItem des + pre + med + cla
            End If
              adodcprod.Recordset.MoveNext
            Loop
            adodcprod.Recordset.MoveFirst
    Call asigna
If adodcprod.Recordset!OFERTADO = True Or (adodcprod.Recordset!MEDMAY > 0) Then
   Me.OFERTADO.Enabled = True
   Me.OFERTADO.Visible = True
   'DESACTIVAR TODOS LOS CONTROLES
   txtprecosto.Enabled = False
   mskcostopza.Enabled = False
   mskcostocaja.Enabled = False
   txtppca.Enabled = False
   txtpnoca.Enabled = False
   For i = 0 To 3
       Txtcar(i).Enabled = False
       Txtdes(i).Enabled = False
       Txtes(i).Enabled = False
       Mskprecio(i).Enabled = False
       Mskprecioa(i).Enabled = False
   Next
  cmdGrabar.Enabled = False
  mediomay.Enabled = False
  cmdGrabar.Visible = False
  mediomay.Visible = False
End If

If Nivel = "O" Or Nivel = "I" Or Nivel = "T" Then
   'Me.cmdGrabar.Enabled = Mid(cSucursal, 1, 2) = "16"
   Me.mediomay.Enabled = True
   Me.cmdGrabar.Enabled = True
   'por si las dudas
Else
   Me.cmdGrabar.Enabled = False
   Me.mediomay.Enabled = False
End If
Call deshabilita_productos
'OTRA MODIFICACION PARA TRABAJAR EN REPLICA LAS ESCALAS
'EN EL CASO DE QUE SEA BODEGA
'If tipotienda <> "B" Then
If tipotienda <> 2 And Sql Then
    For i = 1 To 3
        Txtes(i).Enabled = False
        Mskprecio(i).Enabled = False
    Next
    If adodcprod.Recordset!PAQUETES > 1 Then
       pzasmay = Round(adodcprod.Recordset!PAQUETES / 2 + 0.1)
       Label29(5).Caption = "1/2 Caja Bod. [" & pzasmay & "] Pzas."
       Label29(4).Caption = "1/2 caja Env. [" & pzasmay & "] Pzas."
    Else
       ' Label29(4).Caption = "Medio mayoreo"
    End If
End If
If tipotienda = 4 Then
    For N = 0 To 5
       Txtes(N).Enabled = False
       Mskprecio(N).Enabled = False
    Next
    Txtes(3).Enabled = True
    Mskprecio(3).Enabled = True
End If
'Obtiene el costo de productos
nValAnt = Val(mskcostocaja.Text)
Call calculo
'validatasa Txtcar(2).Text, Txtcar(3).Text
Txtdepto.Text = adodcprod.Recordset!tasaieps
If Sql Then
    Set adorsTemp = New ADODB.Recordset
    adorsTemp.Open "SELECT consec,descripc,str(paquetes) + ' x ' + ltrim(str(contenid,10,3)) + ' ' + medida as medida FROM inventario,cambpre,tfproduc WHERE inprod = consec AND consec = producto and producto = inprod AND modificado = 0 AND incant <= invcamb", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not adorsTemp.EOF
        MsgBox "El Producto " & adorsTemp!DESCRIPC & " " & adorsTemp!medida & " esta registrado para cambio de precio", vbExclamation, "Alerta para cambio de precio"
        adorsTemp.MoveNext
    Wend
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub deshabilita_productos()
'los siguientes productos estan desactivados por default y se deben hacer a mano
clavet = Trim(txtclprod.Text)
Select Case clavet
   Case "1007991"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1007992"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1007993"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1007994"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1007995"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1008833"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1008834"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1008835"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1008836"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1008837"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1008533"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1014960"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
   Case "1007991"
        cmdGrabar.Enabled = False
        mediomay.Enabled = False
End Select
'fin de los productos
End Sub

Private Sub txtcajas_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
    Txtpzaxca.SetFocus
End If

Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
'If OFERTADO.Caption = "OFERTADO" Then
'   Exit Sub
'End If
'If nivel = "C" Or nivel = "T" Then
'   If Not VALIDAPRECIOS Then
'           MsgBox "Es necesario Corregir los Precios...", vbCritical
'          Cancel = True
'  End If
'End If
End Sub

Private Sub gridmm_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 2 Then
    'se debe hacer la actualizacion del precio base
    paqt = Adomm.Recordset!PAQUETES
    clave = Adomm.Recordset!CONSEC
    CADENA = "update tfproduc set costocaj = " & (Mskprecio(0).Text * 0.9) * paqt & "  where consec = '" & Trim(clave) & "'"
    'MsgBox cadena
    cn.Execute CADENA
    MsgBox "se ha actualizado el precio"
End If
End Sub

Private Sub Label34_Click()

End Sub

Private Sub mediomay_Click()
If Nivel <> "O" Then
   Me.cmdGrabar.Enabled = False
End If
'DESPLEGAR EN UN GRID LOS MEDIOS MAY QUE TENGA
'SE ABRE EL COMPONENTE CADA QUE SE DE ESTA OPCION
Adomm.CursorType = adOpenKeyset
Adomm.LockType = adLockOptimistic
Adomm.CommandType = adCmdText
Adomm.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
'CUAL ES EL PRODUCTO BASE
clave = Trim(txtclprod.Text)
Adomm.RecordSource = "select activo,CONSEC,DESCRIPC, PAQUETES, costocaj as precio1 from tfproduc,preprod WHERE consec = preclave and  MEDMAY = " & clave & " ORDER BY descripc"
Adomm.Refresh
frmmm.Visible = True
frmmm.Enabled = True
End Sub



Private Sub Mskprecio_GotFocus(Index As Integer)
 antprecio = Val(Mskprecio(Index).Text)
 Mskprecio(Index).SelStart = 0
 Mskprecio(Index).SelLength = Len(Mskprecio(Index).Text)
End Sub

Private Sub Mskprecio_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   SendKeys "+{tab}"
End If
If KeyCode = 40 Then
   'SendKeys vbTab
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub Mskprecio_LostFocus(Index As Integer)
On Error GoTo Error:
Dim npasoprec As Double
    If antprecio <> Val(Mskprecio(Index).Text) Then
       If Val(Mskpreprom.Text) > 0 Then
          npasoprec = Val(Mskpreprom.Text)
       Else
          npasoprec = Val(txtpreciodes.Text)
       End If
       'TENIA 2
       If Index > 0 And Index <> 4 And Index <> 5 Then  'Que no sea la primera escala y medio mayoreo
          Txtes(Index).Text = (Round(((Val(Mskprecio(Index).Text) * 100) / npasoprec) - 100, 3))
       Else
          If Index = 4 Or Index = 5 Then    'Medio Mayoreo
             Valor = (Round((((Val(Mskprecio(Index).Text)) * 100) / npasoprec) - 100, 3))
          Else  'Autoservicio
             Valor = (Round((((Val(Mskprecio(Index).Text) * IIf(Not IsNull(adodcprod.Recordset!PAQUETES), adodcprod.Recordset!PAQUETES, 1)) * 100) / npasoprec) - 100, 3))
          End If
          Txtes(Index).Text = Valor
       End If
       Txtes(Index).Refresh
       'OTRA MODIFICACION POR SI TIENEN QUE CAMBIAR EL PRECIO
       'ANTES DE MANDAR ACTUALIZAR PRECIOS SE DEBEN ACTUALIZAR SEGUN EL FACTOR LAS ESCALAS
       'If tipotienda = 1 And chkmanual.Value = 0 Then
       If tipotienda = 1 Or tipotienda = 4 And chkmanual.Value = 0 Then
          Call REPLICAESCALAS
       End If
       If Index = 3 And chkmanual.Value = 1 Then
          ESCALA npasoprec, 9   'Cuando se actualiza la escala 3 en precios manuales no se actualizan las demas escalas
       Else
          ESCALA npasoprec, Index
       End If
    End If
    If Index = 3 Then
       'Me.txtprecosto.SetFocus
       'Me.cmdGrabar.SetFocus
    End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtautoriza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Fraautoriza.Enabled = False
    Fraautoriza.Visible = False
    If txtautoriza = "MODI12" Then
        'Call calculo
        'Call REPLICAESCALAS
        'ESCALA costotal
        MsgBox "Contraseña Correcta, Presione el boton Grabar ", vbInformation
        puedegrabar = True
    Else
       MsgBox "Ingrese la Escala correcta", vbCritical
       Txtes(0).SetFocus
    End If
End If
End Sub

Private Sub Txtcar_GotFocus(Index As Integer)
   Txtcar(Index).SelStart = 0
   Txtcar(Index).SelLength = Len(Txtcar(Index).Text)
   nValAnt = Txtcar(Index).Text
End Sub

Private Sub Txtcar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   SendKeys "+{tab}"
End If
If KeyCode = 40 Then
   'SendKeys vbTab
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub Txtcar_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub

Private Sub Txtcar_LostFocus(Index As Integer)
On Error GoTo Error:
 If Trim(Txtcar(Index).Text) = "" Then Txtcar(Index).Text = 0
 Call calculo
 If nValAnt <> Txtcar(Index).Text Then
    Call REPLICAESCALAS
    ESCALA costotal, 3
 End If
 'Regresa el depto. al que pertenece el producto en base al Iva y IEPS
 If Index = 2 Or Index = 3 Then validatasa Txtcar(2).Text, Txtcar(3).Text
 Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtclprod_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
    If Trim(txtclprod.Text) <> "" Then
        If producto(txtclprod.Text) Then
            Call asigna
        Else
            cmdnuevo.Visible = True
            cmdsalir.Visible = False
            txtclprod.Locked = True
            cmbprod.Visible = True
            cmdGrabar.Visible = False
            cmbprod.SetFocus
        End If
    Else
       cmdnuevo.Visible = False
       cmdsalir.Visible = False
       txtclprod.Locked = True
       cmbprod.Visible = True
       cmdGrabar.Visible = False
       cmbprod.SetFocus
    End If
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub
Sub asigna()
On Error GoTo Error:
    'MsgBox adodcprod.Recordset.RecordCount
    prov.Caption = adodcprod.Recordset!claprove
    lblProducto.Caption = adodcprod.Recordset!DESCRIPC
    lblProv.Caption = adodcprod.Recordset!NOMPROVE
    lblfamilia.Caption = adodcprod.Recordset!fdescrip
    Lbldepto.Caption = adodcprod.Recordset!depdescrip
    Lbllinea.Caption = adodcprod.Recordset!sfdescrip
    lblcodbarra.Caption = IIf(Not IsNull(adodcprod.Recordset!barraspza), adodcprod.Recordset!barraspza, 0)
    lblpres.Caption = Trim(Str(adodcprod.Recordset!PAQUETES)) + " X  " + Trim(Str(adodcprod.Recordset!Contenid)) + " " + adodcprod.Recordset!medida
    txtclprod.Text = adodcprod.Recordset!CONSEC
    npzaxcaj = IIf(Not IsNull(adodcprod.Recordset!PAQUETES), adodcprod.Recordset!PAQUETES, 1)
    txtpnoca.Text = IIf(Not IsNull(adodcprod.Recordset!cajas), adodcprod.Recordset!cajas, 0)
    txtppca.Text = IIf(Not IsNull(adodcprod.Recordset!encajas), adodcprod.Recordset!encajas, 0)
    'Igualo observaciones y descuentos financieros
    For N = 1 To 5
        txtDectoFin(N - 1).Text = IIf(IsNull(adodcprod.Recordset.Fields("dectofin" & CStr(N)).Value), "", adodcprod.Recordset.Fields("dectofin" & CStr(N)).Value)
        txtobserva(N - 1).Text = IIf(IsNull(adodcprod.Recordset.Fields("observa" & CStr(N)).Value), "", adodcprod.Recordset.Fields("observa" & CStr(N)).Value)
    Next
    'ESTAN BIEN AUNQUE CON NOMBRES INVERSOS
    'ES EL PRECIO DE LISTA
    txtprecosto.Text = adodcprod.Recordset!costocaj
    'CHECAR ESTO
    
    txtpreciodes.Text = adodcprod.Recordset!PRECOSTO
    Frame2.Visible = True
    'ES EL COSTO
    Me.mskcostocaja = adodcprod.Recordset!PRECOSTO
    Me.mskcostopza = adodcprod.Recordset!COSTOPAQ
    If adodcprod.Recordset!manual = "1" Then
        Me.chkmanual.Value = 1
    Else
        Me.chkmanual.Value = 0
    End If
    Call llenagrid
    'SE DEBE DETECTAR SI ES POR PRIMERA VEZ QUE SE ENTRA
    'Call calculo
Exit Sub
Error:
MsgBox Err.Description

End Sub
Private Function producto(tnClave As String) As Boolean
On Error GoTo Error:
Dim valor1 As Integer
Dim Valor As Integer
Dim strclave As String
Dim strconsec As String
Dim valor2 As Integer
Dim i As Integer
For i = 0 To cmbprod.ListCount - 1
    cmbprod.ListIndex = i
    valor1 = InStr(1, cmbprod.List(cmbprod.ListIndex), "[")
    strconsec = Mid(cmbprod.List(cmbprod.ListIndex), valor1 + 1, 10)
    If Trim(strconsec) = Trim(tnClave) Then
        Valor = InStr(1, cmbprod.List(cmbprod.ListIndex), "\")
        strclave = Mid(cmbprod.List(cmbprod.ListIndex), Valor + 1, Len(cmbprod.List(cmbprod.ListIndex)) - Valor - 1)
        adodcprod.Recordset.Bookmark = Val(strclave)
        producto = True
        i = cmbprod.ListCount
    Else
       producto = False
    End If
Next

Exit Function
Error:
MsgBox Err.Description
End Function

Private Sub txtclprod_LostFocus()
On Error GoTo Error:

'    If Trim(txtclprod.Text) <> "" Then
 '       lsibusca = False
  '      If producto(txtclprod.Text) Then
   '         Call asigna
    '        lsibusca = False
     '        cmdnuevo.Visible = True
      '       cmdsalir.Visible = True
       '      txtclprod.Locked = True
        '     cmbprod.Visible = False
        '     cmdGrabar.Visible = True
        'Else
         '    lsibusca = True
          '   cmdnuevo.Visible = False
          '   cmdsalir.Visible = False
          '   txtclprod.Locked = True
          '   cmbprod.Visible = True
          '   cmdGrabar.Visible = False
          '   cmbprod.SetFocus
        'End If
    'Else
'        lsibusca = True
 '           cmdnuevo.Visible = False
  '           cmdsalir.Visible = False
   '          txtclprod.Locked = True
    '         cmbprod.Visible = True
     '        cmdGrabar.Visible = False
      '       cmbprod.SetFocus
    'End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub txtcontra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cmdConAce_Click
End If

End Sub


Private Sub Txtdepto_GotFocus()
  Txtdepto.SelStart = 0
  Txtdepto.SelLength = Len(Txtdepto.Text)
End Sub

Private Sub TxtDepto_LostFocus()
Dim RST As ADODB.Recordset
Set RST = New ADODB.Recordset
RST.Open "SELECT * FROM tasaieps WHERE depto = " & Me.Txtdepto.Text, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RST.BOF And RST.EOF Then
   MsgBox "Este departamento no existe", vbCritical, "No existe dePto."
   Txtdepto.SetFocus
Else
    Txtcar(2).Text = RST!ieps
    Txtcar(3).Text = RST!iva
    Call calculo
    Call REPLICAESCALAS
    ESCALA costotal, 3
End If
RST.Close
Set RST = Nothing
End Sub

Private Sub Txtdes_GotFocus(Index As Integer)
   Txtdes(Index).SelStart = 0
   Txtdes(Index).SelLength = Len(Txtdes(Index).Text)
   nValAnt = Txtdes(Index).Text
End Sub

Private Sub Txtdes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   SendKeys "+{tab}"
End If
If KeyCode = 40 Then
   'SendKeys vbTab
   keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
End If
End Sub

Private Sub Txtdes_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub


Private Sub Txtdes_LostFocus(Index As Integer)
On Error GoTo Error:
'If chkmanual.Value = 0 Then ESCALA costotal, 0
 If Trim(Txtcar(Index).Text) = "" Then Txtcar(Index).Text = 0
 Call calculo
 If nValAnt <> Txtdes(Index).Text Then
    Call REPLICAESCALAS
    ESCALA costotal, 3
 End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtes_GotFocus(Index As Integer)
 antprecio = Val(Txtes(Index).Text)
 Txtes(Index).SelStart = 0
 Txtes(Index).SelLength = Len(Txtes(Index).Text)
End Sub

Private Sub Txtes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   SendKeys "+{tab}"
End If
If KeyCode = 40 Then
   'SendKeys vbTab
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub Txtes_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Error:
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
    SendKeys "{BACKSPACE}"
    Exit Sub
End If
If KeyAscii = 13 And antprecio <> Txtes(Index).Text Then
'PRIMER CAMBIO PARA QUE SOLO SE ACTUALIZE LA PRIMERA ESCALA
'DE EN CASCADA LAS DEMAS ESCALAS
  If Index = 3 Then
     Call calculo
     If (tipotienda = 1 Or tipotienda = 4) And chkmanual.Value = 0 Then
        Call REPLICAESCALAS
     End If
     If chkmanual.Value = 1 And Index = 3 Then
        ESCALA costotal, 9, True       'Para indicar que el cambio es por la escala
     Else
        ESCALA costotal, Index
     End If
  End If
  'Si es mayoreo o en precios manuales
  If tipotienda = 4 Or chkmanual.Value = 1 Then
     'If Index = 1 Or Index = 2 Or Index = 3 Then
        Call calculo
        ESCALA costotal, Index, True  'Para indicar que el cambio es por la escala
     'End If
  End If
  If Index = 3 Then
     'Me.txtprecosto.SetFocus
     If Sql Then Me.cmdGrabar.SetFocus
  End If
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub REPLICAESCALAS()
If tipotienda = 1 Then
   escalabase = Round(Txtes(0).Text, 3)
   escala2 = Round(escalabase * 0.7, 3)
   escala3 = Round(escalabase * 0.5, 3)
   escala4 = Round(escalabase * 0.3, 3)
   'AHORA SE APLICAN EN LA PANTALLA
   Txtes(0).Text = escalabase 'Escala pieza
   Txtes(1).Text = escala2    'Mayoreo autoservicio
   Txtes(2).Text = escala3    'Mayoreo domicilio
   Txtes(3).Text = escala4    'Mayoreo Bodega
   Txtes(4).Text = IIf(adodcprod.Recordset!PAQUETES >= 10, (escalabase + escala2) / 2, escalabase)   'Escala para Medio mayoreo
Else
   escalabase = Round(Txtes(3).Text, 3)
   escala2 = Round(escalabase + 2, 3)
   escala3 = Round(escalabase + 4, 3)
   escala4 = Round(escalabase + 6, 3)
   escala5 = Round(escalabase + 8, 3)
   escala6 = Round(escalabase + 12, 3)
   'AHORA SE APLICAN EN LA PANTALLA
   Txtes(3).Text = escalabase 'Mayoreo bodega
   Txtes(2).Text = escala2    'Mayoreo Intermedio
   Txtes(1).Text = escala3    'Mayoreo Envío crédito
   Txtes(5).Text = escala4    'Medio Mayoreo Bodega
   Txtes(4).Text = escala5    'Medio Mayoreo Envío
   Txtes(0).Text = escala6    'Pieza Autoservicio
End If
End Sub

Private Sub Txtes_LostFocus(Index As Integer)
On Error GoTo Error:
Txtes_KeyPress Index, 13
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtfecfin_GotFocus()
On Error GoTo Error:
If Trim(txtfecfin.Text) <> "" Then
   Cal1.Value = txtfecfin.Text
Else
   Cal1.Value = date
End If
Cal1.Top = 4680
Cal1.Visible = True
Cal1.SetFocus
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtfecini_GotFocus()
On Error GoTo Error:
If Trim(txtfecini.Text) <> "" Then
    Cal1.Value = txtfecini.Text
Else
    Cal1.Value = date
End If
Cal1.Top = 3600
Cal1.Visible = True
Cal1.SetFocus
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Txtplazo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub

Private Sub txtpnoca_GotFocus()
txtpnoca.SelStart = 0
txtpnoca.SelLength = Len(txtpnoca.Text)
nValAnt = txtpnoca.Text
End Sub

Private Sub txtpnoca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   SendKeys "+{tab}"
End If
If KeyCode = 40 Then
   'SendKeys vbTab
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub txtpnoca_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
    txtppca.SetFocus
    Call calculo
    ESCALA costotal, 0
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Sub llenagrid()
On Error GoTo Error:
    For i = 0 To 6
    Txtdes(i).Text = 0
    Next
    
    For i = 0 To 6
    Txtcar(i).Text = 0
    Next
    
    For i = 0 To 3
        Mskprecio(i).Text = 0
        Mskprecioa(i).Text = 0
        Txtes(i).Text = 0
    Next
    
    Adodccargo.CommandType = adCmdText
    Adodccargo.CursorType = adOpenKeyset
    Adodccargo.LockType = adLockOptimistic
    Adodccargo.ConnectionString = cCadConex
    Adodccargo.RecordSource = "select * from cargos where caprod = '" & Trim(txtclprod.Text) & "'"
    Adodccargo.Refresh
   
   If Adodccargo.Recordset.RecordCount < 1 Then
       cn.Execute "insert into cargos (caprod,cargo1,cargo2,ieps,iva,cargo_efectivo,flete_efectivo,maniobras) VALUES ( '" & Trim(txtclprod.Text) & "',0,0,0,0,0,0,0)"
       Adodccargo.Refresh
   Else
       Txtcar(0).Text = IIf(Not IsNull(Adodccargo.Recordset!cargo1), Adodccargo.Recordset!cargo1, 0)
       Txtcar(1).Text = IIf(Not IsNull(Adodccargo.Recordset!cargo2), Adodccargo.Recordset!cargo2, 0)
       Txtcar(2).Text = IIf(Not IsNull(Adodccargo.Recordset!ieps), Adodccargo.Recordset!ieps, 0)
       Txtcar(3).Text = IIf(Not IsNull(Adodccargo.Recordset!iva), Adodccargo.Recordset!iva, 0)
       Txtcar(4).Text = IIf(Not IsNull(Adodccargo.Recordset!cargo_efectivo), Adodccargo.Recordset!cargo_efectivo, 0)
       Txtcar(5).Text = IIf(Not IsNull(Adodccargo.Recordset!flete_efectivo), Adodccargo.Recordset!flete_efectivo, 0)
       Txtcar(6).Text = IIf(Not IsNull(Adodccargo.Recordset!maniobras), Adodccargo.Recordset!maniobras, 0)
   End If
    
    'abrir tabla de descuentos
    adodcdes.CommandType = adCmdText
    adodcdes.CursorType = adOpenKeyset
    adodcdes.LockType = adLockOptimistic
    adodcdes.ConnectionString = cCadConex
    adodcdes.RecordSource = "select * from descuentos   where deprod = '" & Trim(txtclprod.Text) & "'"
    adodcdes.Refresh
   
   If adodcdes.Recordset.RecordCount < 1 Then
       cn.Execute "insert into descuentos (deprod,decto1,decto2,decto3,dectooferta,dectofinanciero,dectoefectivo) VALUES ( '" & Trim(txtclprod.Text) & "',0,0,0,0,0,0);"
       adodcdes.Refresh
   Else
       Txtdes(0).Text = IIf(Not IsNull(adodcdes.Recordset!decto1), adodcdes.Recordset!decto1, 0)
       Txtdes(1).Text = IIf(Not IsNull(adodcdes.Recordset!decto2), adodcdes.Recordset!decto2, 0)
       Txtdes(2).Text = IIf(Not IsNull(adodcdes.Recordset!decto3), adodcdes.Recordset!decto3, 0)
       Txtdes(3).Text = IIf(Not IsNull(adodcdes.Recordset!dectoOferta), adodcdes.Recordset!dectoOferta, 0)
       Txtdes(4).Text = IIf(Not IsNull(adodcdes.Recordset!dectoFinanciero), adodcdes.Recordset!dectoFinanciero, 0)
       Txtdes(5).Text = IIf(Not IsNull(adodcdes.Recordset!dectoefectivo), adodcdes.Recordset!dectoefectivo, 0)
       Txtdes(6).Text = IIf(Not IsNull(adodcdes.Recordset!decto5), adodcdes.Recordset!decto5, 0)
'       txtfecini.Text = IIf(Not IsNull(adodcdes.Recordset!defecini), adodcdes.Recordset!defecini, "")
'       txtfecfin.Text = IIf(Not IsNull(adodcdes.Recordset!defecfin), adodcdes.Recordset!defecfin, "")
   End If
    'abrir tabla de precios
    adodcprecio.CommandType = adCmdText
    adodcprecio.CursorType = adOpenKeyset
    adodcprecio.LockType = adLockOptimistic
    adodcprecio.ConnectionString = cCadConex
    adodcprecio.RecordSource = "select * from preprod where preclave = '" & Trim(txtclprod.Text) & "'"
    adodcprecio.Refresh
    If adodcprecio.Recordset.RecordCount < 1 Then
        cn.Execute "insert into preprod (preclave,precio1,precio2,precio3,precio4,preciomm,premaydomcre,premaybodcred) VALUES ( '" & Trim(txtclprod.Text) & "',0,0,0,0,0,0,0)"
        adodcprecio.Refresh
        'Lblfechact.Caption = "Fecha Ult. Act.  " & Date
    Else
       Mskprecio(0).Text = IIf(Not IsNull(adodcprecio.Recordset!precio1), adodcprecio.Recordset!precio1, 0)
       Mskprecio(1).Text = IIf(Not IsNull(adodcprecio.Recordset!PRECIO2), adodcprecio.Recordset!PRECIO2, 0)
       Mskprecio(2).Text = IIf(Not IsNull(adodcprecio.Recordset!PRECIO3), adodcprecio.Recordset!PRECIO3, 0)
       Mskprecio(3).Text = IIf(Not IsNull(adodcprecio.Recordset!precio4), adodcprecio.Recordset!precio4, 0)
       Mskprecio(4).Text = IIf(Not IsNull(adodcprecio.Recordset!precio5), adodcprecio.Recordset!precio5, 0)
'       If Sql Then Mskprecio(5).Text = IIf(Not IsNull(adodcprecio.Recordset!precio6), adodcprecio.Recordset!precio6, 0)
       Mskprecio(5).Text = IIf(Not IsNull(adodcprecio.Recordset!precio6), adodcprecio.Recordset!precio6, 0)
       Mskprecioa(0).Text = IIf(Not IsNull(adodcprecio.Recordset!precio1ant), adodcprecio.Recordset!precio1ant, 0)
       Mskprecioa(1).Text = IIf(Not IsNull(adodcprecio.Recordset!precio2ant), adodcprecio.Recordset!precio2ant, 0)
       Mskprecioa(2).Text = IIf(Not IsNull(adodcprecio.Recordset!precio3ant), adodcprecio.Recordset!precio3ant, 0)
       Mskprecioa(3).Text = IIf(Not IsNull(adodcprecio.Recordset!precio4ant), adodcprecio.Recordset!precio4ant, 0)
       Mskprecioa(4).Text = IIf(Not IsNull(adodcprecio.Recordset!precio5ant), adodcprecio.Recordset!precio5ant, 0)
'       If Sql Then Mskprecioa(5).Text = IIf(Not IsNull(adodcprecio.Recordset!precio6ant), adodcprecio.Recordset!precio6ant, 0)
       Mskprecioa(5).Text = IIf(Not IsNull(adodcprecio.Recordset!precio6ant), adodcprecio.Recordset!precio6ant, 0)
       'AQUI SE DEBE PONER LA FECHA DE ACTUALIZACION Y USUARIO
      'DEBE SER LA FECHA PERO DE ACTUALIZACION DEL PRECIO Y TAMBIEN EL USUARIO QUE MODIFICO EL PRECIO
      Lblfechact.Caption = "Fecha Ult. Act.   " & adodcprecio.Recordset!fechaact
      Me.txtusuario.Caption = "Modifico: " & adodcprecio.Recordset!prusuario
    End If
    'abrir tabla de escalas
    Adoescala.CommandType = adCmdText
    Adoescala.CursorType = adOpenKeyset
    Adoescala.LockType = adLockOptimistic
    Adoescala.ConnectionString = cCadConex
    Adoescala.RecordSource = "select * from margen where producto = '" & Trim(txtclprod.Text) & "'"
    Adoescala.Refresh
   
   If Adoescala.Recordset.RecordCount < 1 Then
       cn.Execute "insert into margen (producto,escala1,escala2,escala3,escala4,mediomayoreo,maydomcred,maybodcred) VALUES ( '" & Trim(txtclprod.Text) & "',0,0,0,0,0,0,0)"
       Adoescala.Refresh
   Else
       Txtes(0).Text = IIf(Not IsNull(Adoescala.Recordset!escala1), Adoescala.Recordset!escala1, 0)
       Txtes(1).Text = IIf(Not IsNull(Adoescala.Recordset!escala2), Adoescala.Recordset!escala2, 0)
       Txtes(2).Text = IIf(Not IsNull(Adoescala.Recordset!escala3), Adoescala.Recordset!escala3, 0)
       Txtes(3).Text = IIf(Not IsNull(Adoescala.Recordset!escala4), Adoescala.Recordset!escala4, 0)
       Txtes(4).Text = IIf(Not IsNull(Adoescala.Recordset!escala5), Adoescala.Recordset!escala5, 0)
       If Sql Then Txtes(5).Text = IIf(Not IsNull(Adoescala.Recordset!escala6), Adoescala.Recordset!escala6, 0)
       Txtes(5).Text = IIf(Not IsNull(Adoescala.Recordset!escala6), Adoescala.Recordset!escala6, 0)
   End If
   If adodcprod.Recordset!PAQUETES > 0 Then
    For i = 1 To 3
        Mskpreciopza(i).Text = Mskprecio(i).Text / adodcprod.Recordset!PAQUETES
    Next
    pzasmay = Round(adodcprod.Recordset!PAQUETES / 2 + 0.1)
    Mskpreciopza(0).Text = Mskprecio(0).Text
    Mskpreciopza(4).Text = IIf(pzasmay >= 10, IIf(Trim(Mskprecio(4).Text) = "", 0, Mskprecio(4).Text) / pzasmay, Mskprecio(4).Text)
    Mskpreciopza(5).Text = IIf(pzasmay >= 10, IIf(Trim(Mskprecio(5).Text) = "", 0, Mskprecio(5).Text) / pzasmay, Mskprecio(5).Text)
    End If
Exit Sub
Error:
MsgBox Err.Description
    
End Sub

Private Sub Txtpzaxca_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
    txtpnoca.SetFocus
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub calculo()
'CALCULO DE LOS PRECIOS
On Error GoTo Error:
Dim npreciodes As Double
Dim npreciocar As Double
Dim nprecio As Double
Dim npreciopaso As Double
Dim COMPARA1 As Double
Dim COMPARA2 As Double
Dim i As Integer
If adodcprod.Recordset!OFERTADO = 0 Then
    nprecio = Val(txtprecosto.Text)
    npreciopaso = Val(txtprecosto.Text)
    'calcula cargos %
    For i = 0 To 3
        nprecio = nprecio + (nprecio * (Val(Txtcar(i).Text) / 100))
        If i > 1 Then 'calculo para costo del descprod
           npreciopaso = npreciopaso + (npreciopaso * (Val(Txtcar(i).Text) / 100))
        End If
    Next
    'cargo efectivo
    If ZONA = "OAX" Then
       nprecio = nprecio + Val(Txtcar(4).Text)
    End If

    npreciocar = Round(nprecio, 2)
    Txtpreciocargo.Text = npreciocar
    costodescprod = Round(npreciopaso, 2)
    npreciopaso = 0
    'calcula descuentos %
    nprecio = npreciocar
    For i = 0 To 3
        nprecio = nprecio - (nprecio * (Val(Txtdes(i).Text) / 100))
        costodescprod = costodescprod - (costodescprod * (Val(Txtdes(i).Text) / 100))
    Next
    'el descuento numero 5
    'MsgBox Txtdes(6).Text
    nprecio = nprecio - (nprecio * (Val(Txtdes(6).Text) / 100))
    costodescprod = costodescprod - (costodescprod * (Val(Txtdes(6).Text) / 100))
    
    'descuento efectivo
    'MsgBox Txtdes(4).Text
    nprecio = nprecio - Val(Txtdes(4).Text)
    costodescprod = costodescprod - Val(Txtdes(4).Text)

    'descuento financiero
    'MsgBox Txtdes(5).Text
    nprecio = nprecio - (nprecio * (Val(Txtdes(5).Text) / 100))
    costodescprod = costodescprod - (costodescprod * (Val(Txtdes(5).Text) / 100))
    PRECOSTO = nprecio

    If ZONA = "CHS" Then
       nprecio = nprecio + Val(Txtcar(4).Text)
    End If


    'cargo por flete
    nprecio = nprecio + Val(Txtcar(5).Text)
    nprecio = nprecio + Val(Txtcar(6).Text)

    npreciodes = Round(nprecio, 2)
    txtpreciodes.Text = npreciodes
    nprecio = npreciodes
    ' CALCULO PARA LAS PROMOCIONES
    If Trim(txtppca.Text) <> "" And Trim(txtpnoca.Text) <> "" Then
      If Val(txtppca.Text) > 0 And Val(txtpnoca.Text) > 0 Then
        nprecio = (nprecio * Val(txtppca.Text)) / (Val(txtppca.Text) + Val(txtpnoca.Text))
        nprecio = Round(nprecio, 2)
        Mskpreprom.Text = nprecio
      Else
        Mskpreprom.Text = 0
      End If
    Else
    txtppca.Text = 0
    txtpnoca.Text = 0
    End If
        'chiapas

    'If mskcostocaja.Text <> nprecio And chkmanual.Value = 1 Then
    '   MsgBox "SE MODIFICARAN SOLAMENTE PRECIOS DE COSTO, LOS PRECIOS DE VENTA PERMANECEN FIJOS PORQUE ESTAN EN MODO MANUALMENTE", vbInformation, "Precios manualmente"
        'End If
    Me.mskcostocaja.Text = nprecio
    Me.mskcostopza.Text = nprecio / adodcprod.Recordset!PAQUETES
    PRECOSTO = nprecio
    costotal = Round(nprecio, 2)
Else
   MsgBox "No se Puede hacer el calculo sobre un producto ofertado..."
End If
Exit Sub
Error:
 MsgBox Err.Description
End Sub

'Calcula el precio, y lo redondea cuando es menor o igual a 20 pesosa 10 centavos en caso contario
'lo redondea a 50 centavos esto se hace siempre y cuando no se haga mediante la opcion preciso manuales.
'ericmag
Sub ESCALA(nprecio As Double, ESCALA As Integer, Optional porEscala As Boolean)
Dim COMPARA1 As Currency
Dim COMPARA2 As Currency
On Error GoTo Error:
 npreciopaso = (nprecio + (nprecio * Val(Txtes(0).Text) / 100)) / npzaxcaj
 preciobase = npreciopaso
 COMPARA1 = Int(npreciopaso)
 COMPARA2 = npreciopaso - COMPARA1
 'compara2 = 0.55
 ' ESTA PARTE DEBE SER POR CADA 10 CENTAVOS

If ESCALA = 3 Then   'Cuando modifican precios de lista, cargos o descuentos
   'If preciobase <= 20 Then
      'If COMPARA2 > 0 Then
      '   If COMPARA2 <= 0.1 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.1
      '   ElseIf COMPARA2 <= 0.2 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.2
      '   ElseIf COMPARA2 <= 0.3 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.3
      '   ElseIf COMPARA2 <= 0.4 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.4
      '   ElseIf COMPARA2 <= 0.5 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.5
      '   ElseIf COMPARA2 <= 0.6 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.6
      '   ElseIf COMPARA2 <= 0.7 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.7
      '   ElseIf COMPARA2 <= 0.8 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.8
      '   ElseIf COMPARA2 <= 0.9 Then
      '      Mskprecio(0).Text = COMPARA1 + 0.9
      '   Else
      '      Mskprecio(0).Text = COMPARA1 + 1
      '   End If
      'Else
      '   Mskprecio(0).Text = COMPARA1 'si el precio es un entero
      'End If
   'Else
        Mskprecio(0).Text = Redondea(COMPARA1, COMPARA2)  'Funcion que redondea a multiplo de 10 centavos
   'End If
   'Para la bodega de mayoreo no se actualizan precios en cascada
   'If tipotienda = 4 Then Exit Sub
   'AQUI ESTA EL PROBLEMA DE COMPATIBILIDAD DE PRECIOS CON EL ANTERIOR
   For i = 1 To 5
      npreciopaso = Round(nprecio + (nprecio * Val(Txtes(i).Text) / 100), 2)
      If tipotienda <> 2 Then
         preciobase = npreciopaso
         'Se hace el calculo para Medio Mayoreo
         'If i = 4 Then
         '   If adodcprod.Recordset!PAQUETES >= 10 Then
         '      escMedMay = (Val(Txtes(0).Text) + Val(Txtes(1).Text)) / 2
         '      npreciopaso = (nprecio + (nprecio * escMedMay / 100)) / npzaxcaj
         '   Else  'Cuando es 1 x 1 se le da el mismo precio de Autoservicio
         '      npreciopaso = (nprecio + (nprecio * Val(Txtes(0).Text) / 100)) / npzaxcaj
         '   End If
         'End If
      End If
      COMPARA1 = Int(npreciopaso)
      COMPARA2 = npreciopaso - COMPARA1
      
      If preciobase <= 20 Then
         If COMPARA2 > 0 Then
            If COMPARA2 <= 0.1 Then
               Mskprecio(i).Text = COMPARA1 + 0.1
            ElseIf COMPARA2 <= 0.2 Then
                Mskprecio(i).Text = COMPARA1 + 0.2
            ElseIf COMPARA2 <= 0.3 Then
                Mskprecio(i).Text = COMPARA1 + 0.3
            ElseIf COMPARA2 <= 0.4 Then
                Mskprecio(i).Text = COMPARA1 + 0.4
            ElseIf COMPARA2 <= 0.5 Then
                Mskprecio(i).Text = COMPARA1 + 0.5
            ElseIf COMPARA2 <= 0.6 Then
                Mskprecio(i).Text = COMPARA1 + 0.6
            ElseIf COMPARA2 <= 0.7 Then
                Mskprecio(i).Text = COMPARA1 + 0.7
            ElseIf COMPARA2 <= 0.8 Then
                Mskprecio(i).Text = COMPARA1 + 0.8
            ElseIf COMPARA2 <= 0.9 Then
                Mskprecio(i).Text = COMPARA1 + 0.9
            Else
                Mskprecio(i).Text = COMPARA1 + 1
            End If
         Else
            Mskprecio(i).Text = COMPARA1 'si el precio es un entero
         End If
      Else
        If adodcprod.Recordset!PAQUETES > 1 And (i = 4 Or i = 5) Then
           COMPARA1 = Int(npreciopaso / 2)
           COMPARA2 = npreciopaso / 2 - COMPARA1
           Mskprecio(i).Text = Redondea(COMPARA1, COMPARA2)  'Funcion que redondea a multiplo de 10 centavos
        Else
           Mskprecio(i).Text = Redondea(COMPARA1, COMPARA2)  'Funcion que redondea a multiplo de 10 centavos
        End If
      End If
   Next
   
Else
   'ACTUALIZA UNA SOLA ESCALA
   If ESCALA = 9 Then  'Cuando se actualiza la escala 0 en precios manuales no se actualizan las demas escalas
      i = 3
   Else
      i = ESCALA
      If i = 0 Then
         npreciopaso = (nprecio + (nprecio * Val(Txtes(0).Text) / 100)) / npzaxcaj
      Else
         npreciopaso = Round(nprecio + (nprecio * Val(Txtes(i).Text) / 100), 2)
      End If
   End If
   If tipotienda <> 2 Then
      preciobase = npreciopaso
      'SAVEMC
      'If i = 2 Then
      '   If adodcprod.Recordset!PAQUETES > 1 Then
      '      escMedMay = (Val(Txtes(0).Text) + Val(Txtes(1).Text)) / 2
      '      npreciopaso = (nprecio + (nprecio * escMedMay / 100)) / npzaxcaj
      '   Else  'Cuando es 1 x 1 se le da el mismo precio de Autoservicio
      '         npreciopaso = (nprecio + (nprecio * Val(Txtes(0).Text) / 100)) / npzaxcaj
      '   End If
      'End If
   End If
   'En precios manuales no se hacen redondeos
   COMPARA1 = Int(npreciopaso)
   COMPARA2 = npreciopaso - COMPARA1
   If preciobase <= 20 Then
      'If compara2 > 0 Then
      If COMPARA2 > 0 And chkmanual.Value = 0 Then   'save
         If COMPARA2 <= 0.1 Then
            Mskprecio(i).Text = COMPARA1 + 0.1
         ElseIf COMPARA2 <= 0.2 Then
            Mskprecio(i).Text = COMPARA1 + 0.2
         ElseIf COMPARA2 <= 0.3 Then
            Mskprecio(i).Text = COMPARA1 + 0.3
         ElseIf COMPARA2 <= 0.4 Then
            Mskprecio(i).Text = COMPARA1 + 0.4
         ElseIf COMPARA2 <= 0.5 Then
            Mskprecio(i).Text = COMPARA1 + 0.5
         ElseIf COMPARA2 <= 0.6 Then
            Mskprecio(i).Text = COMPARA1 + 0.6
         ElseIf COMPARA2 <= 0.7 Then
            Mskprecio(i).Text = COMPARA1 + 0.7
         ElseIf COMPARA2 <= 0.8 Then
            Mskprecio(i).Text = COMPARA1 + 0.8
         ElseIf COMPARA2 <= 0.9 Then
            Mskprecio(i).Text = COMPARA1 + 0.9
         Else
            Mskprecio(i).Text = COMPARA1 + 1
         End If
      Else
         If chkmanual.Value = 0 Then Mskprecio(i).Text = COMPARA1  'si el precio es un entero
      End If
   Else
      'SAVE última modificación para que no agregue el peso en los enteros
      If chkmanual.Value = 0 Then
         Mskprecio(i).Text = Redondea(COMPARA1, COMPARA2)  'Funcion que redondea a multiplo de 10 centavos
      End If
   End If
   If porEscala Then Mskprecio(i).Text = Round(npreciopaso, 3)
End If

pzasmay = Round(adodcprod.Recordset!PAQUETES / 2 + 0.1)
If (ESCALA = 4 Or ESCALA = 5) Then Mskprecio(ESCALA).Text = Val(Mskprecio(ESCALA).Text) / 2
For i = 0 To 5
   If i = 0 Then
      Mskpreciopza(i).Text = Mskprecio(i).Text
   ElseIf i = 4 Or i = 5 Then
      Mskpreciopza(i).Text = Val(Mskprecio(i).Text) / pzasmay
   Else
      Mskpreciopza(i).Text = Mskprecio(i).Text / adodcprod.Recordset!PAQUETES
   End If
Next
Exit Sub
Error:
MsgBox Err.Description

End Sub

'Redondea el precio calculado, todo se redondea a multiplos de 10 centavos
Private Function Redondea(preEntero As Currency, preDecimal As Currency)
If preDecimal = 0 Then
   Redondea = preEntero
Else    'Todo se redondea a multiplos de 5 centavos
  If preDecimal <= 0.05 Then
     Redondea = preEntero + 0.05
  ElseIf preDecimal <= 0.1 Then
     Redondea = preEntero + 0.1
  ElseIf preDecimal <= 0.15 Then
     Redondea = preEntero + 0.15
  ElseIf preDecimal <= 0.19 Then
     Redondea = preEntero + 0.2
  ElseIf preDecimal <= 0.25 Then
     Redondea = preEntero + 0.25
  ElseIf preDecimal <= 0.3 Then
     Redondea = preEntero + 0.3
  ElseIf preDecimal <= 0.35 Then
     Redondea = preEntero + 0.35
  ElseIf preDecimal <= 0.4 Then
     Redondea = preEntero + 0.4
  ElseIf preDecimal <= 0.45 Then
     Redondea = preEntero + 0.45
  ElseIf preDecimal <= 0.5 Then
     Redondea = preEntero + 0.5
  ElseIf preDecimal <= 0.55 Then
     Redondea = preEntero + 0.55
  ElseIf preDecimal <= 0.6 Then
     Redondea = preEntero + 0.6
  ElseIf preDecimal <= 0.65 Then
     Redondea = preEntero + 0.65
  ElseIf preDecimal <= 0.7 Then
     Redondea = preEntero + 0.7
  ElseIf preDecimal <= 0.75 Then
     Redondea = preEntero + 0.75
  ElseIf preDecimal <= 0.8 Then
     Redondea = preEntero + 0.8
  ElseIf preDecimal <= 0.85 Then
     Redondea = preEntero + 0.85
  ElseIf preDecimal <= 0.9 Then
     Redondea = preEntero + 0.9
  ElseIf preDecimal <= 0.95 Then
     Redondea = preEntero + 0.95
  'ElseIf preDecimal <= 0.99 Then
  '   Redondea = preEntero + 0.9
  Else
     Redondea = preEntero + 1
  End If
End If
End Function

Private Sub txtpnoca_LostFocus()
On Error GoTo Error:
If Trim(txtpnoca.Text) <> "" Then
If Int(Val(txtpnoca.Text)) <> Val(txtpnoca.Text) Then txtpnoca.Text = Int(Val(txtpnoca.Text))
    Call calculo
    Call REPLICAESCALAS
    If nValAnt <> txtpnoca.Text Then ESCALA costotal, 3
Else
 txtpnoca.Text = 0
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtppca_GotFocus()
txtppca.SelStart = 0
txtppca.SelLength = Len(txtppca.Text)
nValAnt = txtppca.Text
End Sub

Private Sub txtppca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   SendKeys "+{tab}"
End If
If KeyCode = 40 Then
   'SendKeys vbTab
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub txtppca_LostFocus()
On Error GoTo Error:
If Trim(txtppca.Text) <> "" Then
    If Int(Val(txtppca.Text)) <> Val(txtppca.Text) Then txtppca.Text = Int(Val(txtppca.Text))
    Call calculo
    If nValAnt <> txtppca.Text Then
       ESCALA costotal, 3
       Call REPLICAESCALAS
    End If
Else
 txtppca.Text = 0
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtprecosto_GotFocus()
 txtprecosto.SelStart = 0
 txtprecosto.SelLength = Len(txtprecosto.Text)
 nValAnt = Val(txtprecosto.Text)
End Sub

Private Sub txtprecosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   SendKeys "+{tab}"
End If
If KeyCode = 40 Then
   keybd_event &H9, 0, 0, 0
End If
End Sub

Private Sub txtprecosto_LostFocus()
On Error GoTo Error:
Call calculo
If nValAnt <> txtprecosto.Text Then
  Call REPLICAESCALAS
  ESCALA costotal, 3
End If
Exit Sub
Error:
  MsgBox Err.Description
End Sub
