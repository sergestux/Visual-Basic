VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form fmenu 
   Caption         =   "Servicio a Compras..."
   ClientHeight    =   5715
   ClientLeft      =   1725
   ClientTop       =   2100
   ClientWidth     =   8145
   Icon            =   "menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8145
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoMargen 
      Height          =   330
      Left            =   120
      Top             =   4320
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
      Caption         =   "Margen"
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
   Begin VB.Frame Fraimporta 
      Height          =   4215
      Left            =   2760
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Terminar"
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   18
         Top             =   3360
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar ProBar1 
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Importar"
         Height          =   615
         Index           =   0
         Left            =   960
         Picture         =   "menu.frx":400A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Exportar"
         Height          =   615
         Index           =   1
         Left            =   4560
         Picture         =   "menu.frx":444C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Lbltrans 
         Caption         =   "Realizando Transacción ......"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Frame Frmabrir 
      Caption         =   "Abriendo Catalogos ..."
      Height          =   975
      Left            =   2880
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Data DaoProdDbf 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data DaoProvDbf 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   765
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11775
      Begin VB.CommandButton Cmdaccion 
         Caption         =   "&Utilerias"
         Height          =   855
         Index           =   4
         Left            =   8640
         Picture         =   "menu.frx":488E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Reportes"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Or&ganizar"
         Height          =   855
         Left            =   6240
         Picture         =   "menu.frx":4CD0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Organizar Productos"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmdaccion 
         Caption         =   "&Ofertas"
         Height          =   855
         Index           =   3
         Left            =   5040
         Picture         =   "menu.frx":5112
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Calcular Precios"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   10560
         Picture         =   "menu.frx":541C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir del Sistema"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Cmdaccion 
         Caption         =   "&Reportes"
         Height          =   855
         Index           =   7
         Left            =   7440
         Picture         =   "menu.frx":585E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Reportes"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmdaccion 
         Caption         =   "&Hoja de Cat."
         Height          =   855
         Index           =   6
         Left            =   240
         Picture         =   "menu.frx":5CA0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Hoja de Catalogos "
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmdaccion 
         Caption         =   "Pre&cios"
         Height          =   855
         Index           =   2
         Left            =   3840
         Picture         =   "menu.frx":60E2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Calcular Precios"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmdaccion 
         Caption         =   "&Productos"
         Height          =   855
         Index           =   1
         Left            =   2640
         Picture         =   "menu.frx":63EC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Catalogo de Productos"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmdaccion 
         Caption         =   "Pro&veedores"
         Height          =   855
         Index           =   0
         Left            =   1440
         Picture         =   "menu.frx":66F6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Catalogo de provedores"
         Top             =   240
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   720
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   10560
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin MSComctlLib.StatusBar stbMensajes 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   5370
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3087
            MinWidth        =   3087
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adopreprod 
      Height          =   330
      Left            =   0
      Top             =   3525
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
      Caption         =   "preprod"
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
   Begin MSAdodcLib.Adodc AdoDescuentos 
      Height          =   330
      Left            =   0
      Top             =   3165
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
   Begin MSAdodcLib.Adodc AdoCargos 
      Height          =   330
      Left            =   0
      Top             =   3885
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
   Begin MSAdodcLib.Adodc AdoTfproduc 
      Height          =   375
      Left            =   0
      Top             =   2685
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
      Caption         =   "tfproduc"
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
   Begin MSAdodcLib.Adodc AdoProv 
      Height          =   330
      Left            =   75
      Top             =   1965
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
      Caption         =   "Adoprov"
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
   Begin MSAdodcLib.Adodc AdoProd 
      Height          =   375
      Left            =   0
      Top             =   2325
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
      Caption         =   "Adoprod"
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
   Begin MSAdodcLib.Adodc AdoDescprod 
      Height          =   330
      Left            =   0
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
      Caption         =   "descprod"
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
   Begin VB.Label LblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   11895
   End
   Begin VB.Image ImgTapiz 
      Appearance      =   0  'Flat
      Height          =   5955
      Left            =   0
      Picture         =   "menu.frx":6B38
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8175
   End
   Begin VB.Menu mnuocat 
      Caption         =   "&Catalogos"
      Visible         =   0   'False
      Begin VB.Menu mnucatpro 
         Caption         =   "&Proveedores"
      End
      Begin VB.Menu mnwcatdep 
         Caption         =   "&Organizar Departamentos"
         Begin VB.Menu mnucadepto 
            Caption         =   "&Departamentos"
         End
         Begin VB.Menu mnufam 
            Caption         =   "&Familias"
         End
         Begin VB.Menu mnucalinea 
            Caption         =   "&Lineas"
         End
      End
      Begin VB.Menu mnuprod 
         Caption         =   "P&roductos"
      End
      Begin VB.Menu mnuprec 
         Caption         =   "$ Pre&cios"
      End
      Begin VB.Menu mnuoferta 
         Caption         =   "Ofer&tas"
      End
      Begin VB.Menu mnurep 
         Caption         =   "Repor&tes"
      End
      Begin VB.Menu mnusal 
         Caption         =   "Sa&lir"
      End
   End
End
Attribute VB_Name = "fmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmdaccion_Click(Index As Integer)
Frame1.Visible = False
Select Case Index
Case 0
    lpprov = False
    fprov.Show 1
Case 1
    stbMensajes.SimpleText = "Espere un momento por favor . Cargando Catalogo de Productos ...... "
    stbMensajes.Refresh
    lpprod = False
    frmprod.Show 1
Case 2
   
    stbMensajes.SimpleText = "Espere un momento por favor . Cargando Catalogo de Productos ...... "
    stbMensajes.Refresh
    lpprov = False
    lpprod = False
    fprecios.Show 1
Case 3
    Fofertas.Show 1
Case 4
   utileria
Case 6
   fhojacat.Show 1

Case 7
    Freporte.Show 1
End Select

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub



Private Sub Command1_Click()
PopupMenu mnwcatdep
End Sub

Private Sub Form_Activate()
LblTitulo.Caption = Chr(13) + "SISTEMA INTEGRAL DE INFORMACION" + Chr(13) + "MODULO DE COSTOS"
LblTitulo.Refresh
Frame1.Visible = True
End Sub

Private Sub Form_Load()
Dim mes(11) As Variant

  mes(0) = "Enero": mes(1) = "Febrero": mes(2) = "Marzo": mes(3) = "Abril":  mes(4) = "Mayo": mes(5) = "Junio": mes(6) = "Julio"
  mes(7) = "Agosto": mes(8) = "Septiembre":  mes(9) = "Octubre": mes(10) = "Noviembre": mes(11) = "Diciembre"
  
stbMensajes.Panels(1).Text = Space(20) & "Alt + tecla resaltada activa Menú"
stbMensajes.Panels(2).Text = Space(5) & Str(Day(Date)) + " de " + mes(Month(Date) - 1) & " del " & Str(Year(Date))
stbMensajes.Panels(3).Text = Space(5) & Time
Me.Caption = ccaption
Unload frmLogin1
lpasalog = False
End Sub


Private Sub Form_Resize()
On Error Resume Next
stbMensajes.Panels(1).Width = fmenu.Width / 2
stbMensajes.Panels(2).Width = fmenu.Width / 4
stbMensajes.Panels(3).Width = fmenu.Width / 4
ImgTapiz.Width = fmenu.ScaleWidth
ImgTapiz.Height = fmenu.ScaleHeight
Frame1.Width = fmenu.ScaleWidth - 200
End Sub


Private Sub ImgTapiz_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 PopupMenu mnuocat
End Sub

Private Sub mnucadepto_Click()
fdeptos.Show 1

End Sub

Private Sub mnucalinea_Click()
flineas.Show 1

End Sub

Private Sub mnucatpro_Click()
fprov.Show 1

End Sub

Private Sub mnudesca_Click()
frmocat.Show 1

End Sub

Private Sub mnuescalas_Click()
Frmescala.Show 1

End Sub

Private Sub mnufam_Click()
ffamilia.Show 1

End Sub

Private Sub mnuoferta_Click()
Fofertas.Show 1
End Sub

Private Sub mnuprec_Click()
lpprov = False
lpprod = False
fprecios.Show 1

End Sub

Private Sub mnuprod_Click()
stbMensajes.SimpleText = "Espere un momento por favor . Cargando Catalogo de Productos ...... "
stbMensajes.Refresh
lpprod = False
frmprod.Show 1
End Sub

Private Sub mnusal_Click()
Unload Me

End Sub
Sub utileria()

Adoprov.CursorType = adOpenKeyset
Adoprov.LockType = adLockOptimistic
Adoprov.CommandType = adCmdText
Adoprov.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
Adoprov.RecordSource = "select * from catprov"
Adoprov.Refresh

Fraimporta.Visible = True
Fraimporta.Refresh

End Sub
Private Sub Cmdopt_Click(Index As Integer)
Select Case Index
Case 1
    Exporta
Case 0
    Importa
Case 2
    Fraimporta.Visible = False
    Fraimporta.Refresh
    Frame1.Visible = True
    Frame1.Refresh
End Select

End Sub

Sub Exporta()
On Error GoTo error:

'este proceso utiliza 3 dbfs en el directorio c:\paso
'prodnob.dbf es la estructura de productos que se copia al archivo producto
'producto.dbf y productox.dbf se utiliza para que el recordset libere la tabla de productos
'y se pueda utilizar en el proceso de importacion o a la inversa

respsn = MsgBox("  Exportación TOTAL", vbQuestion + vbYesNoCancel, "Utilerias")
Select Case respsn
Case vbYes
    Adoprod.CursorType = adOpenKeyset
    Adoprod.LockType = adLockOptimistic
    Adoprod.CommandType = adCmdText
    Adoprod.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
    Adoprod.RecordSource = "select consec,claprove,descripc,nomcorto,contenid,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja," & _
    "decto1,decto2,decto3,decto4,decto5,flete,maniobras,cargo1,cargo2,cargo3,cargo4,cargo5,financiero,efectivo,descprod.cajas,descprod.encajas," & _
    "precio1,precio2,precio3,precio4,escala1,escala2,escala3,escala4, fecact,tasaieps,peso,ofertado" & _
    " from tfproduc,DESCPROD,preprod,margen where consec = descprod.producto and consec=margen.producto and consec = preclave"
Case vbNo
    Adoprod.CursorType = adOpenKeyset
    Adoprod.LockType = adLockOptimistic
    Adoprod.CommandType = adCmdText
    Adoprod.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
    Adoprod.RecordSource = "select consec,claprove,descripc,nomcorto,contenid,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja," & _
    "decto1,decto2,decto3,decto4,decto5,flete,maniobras,cargo1,cargo2,cargo3,cargo4,cargo5,financiero,efectivo,descprod.cajas,descprod.encajas," & _
    "precio1,precio2,precio3,precio4,escala1,escala2,escala3,escala4, fecact,tasaieps,peso,ofertado" & _
    " from tfproduc,DESCPROD,preprod,margen where consec = descprod.producto and consec=margen.producto and consec = preclave and  tfproduc.actualizado = 1"
Case vbCancel
    Exit Sub
End Select
    Adoprod.Refresh

    Dim v As Integer
    Dim fs

If Adoprod.Recordset.RecordCount > 0 Then
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.copyfile "c:\paso\prodnob.dbf", "c:\paso\producto.dbf", True
    fs.copyfile "c:\paso\provnob.dbf", "c:\paso\catprov.dbf", True
    
    DaoProdDbf.Connect = "dbase III;"
    DaoProdDbf.DatabaseName = "C:\paso"
    DaoProdDbf.RecordsetType = Table
    DaoProdDbf.RecordSource = "producto"
    DaoProdDbf.Refresh
    
    ProBar1.Min = 0
    ProBar1.Max = IIf(Adoprod.Recordset.RecordCount > 0, Adoprod.Recordset.RecordCount, 1)
    Lbltrans.Visible = True
    Lbltrans.Refresh
    ProBar1.Visible = True
    v = 0
    
    Adoprod.Recordset.MoveFirst
    While Not Adoprod.Recordset.EOF
        v = v + 1
        ProBar1.Value = v
            
            concomilla = InStr(1, Adoprod.Recordset!descripc, "'")
            If concomilla > 0 Then
             xdescripc = Mid(Adoprod.Recordset!descripc, 1, concomilla - 1) + "  " + Mid(Adoprod.Recordset!descripc, concomilla + 1, Len(Adoprod.Recordset!descripc))
            Else
             xdescripc = Adoprod.Recordset!descripc
            End If
            
            concomillan = InStr(1, Adoprod.Recordset!Nomcorto, "'")
            If concomillan > 0 Then
             xnomcorto = Mid(Adoprod.Recordset!Nomcorto, 1, concomillan - 1) + "  " + Mid(Adoprod.Recordset!Nomcorto, concomillan + 1, Len(Adoprod.Recordset!Nomcorto))
            Else
             xnomcorto = Adoprod.Recordset!Nomcorto
            End If

    
      Clave = Str(Val(Adoprod.Recordset!consec) - 1000000)
      DaoProdDbf.Recordset.AddNew
      DaoProdDbf.Recordset!consec = Mid(Trim(Clave), 1, 8) ' el campo no cabra en el registros
      DaoProdDbf.Recordset!descripc = IIf(Not IsNull(Adoprod.Recordset!descripc), Mid(Trim(xdescripc), 1, 50), " ")
      DaoProdDbf.Recordset!Nomcorto = IIf(Not IsNull(Adoprod.Recordset!Nomcorto), Mid(xnomcorto, 1, 20), " ") 'el campo es demasiado paqueño para los datos
      DaoProdDbf.Recordset!Contenid = Val(IIf(Not IsNull(Adoprod.Recordset!Contenid), Adoprod.Recordset!Contenid, 0))
      DaoProdDbf.Recordset!fletex = Val(IIf(Not IsNull(Adoprod.Recordset!fletex), Adoprod.Recordset!fletex, 0))
      DaoProdDbf.Recordset!flesub = Val(IIf(Not IsNull(Adoprod.Recordset!flesub), Adoprod.Recordset!flesub, 0))
      DaoProdDbf.Recordset!Medida = IIf(Not IsNull(Adoprod.Recordset!Medida), Adoprod.Recordset!Medida, " ")
      DaoProdDbf.Recordset!Paquetes = Val(IIf(Not IsNull(Adoprod.Recordset!Paquetes), Adoprod.Recordset!Paquetes, 0))
      DaoProdDbf.Recordset!costopaq = Val(IIf(Not IsNull(Adoprod.Recordset!costopaq), Adoprod.Recordset!costopaq, 0))
      DaoProdDbf.Recordset!prelista = Val(IIf(Not IsNull(Adoprod.Recordset!costocaj), Adoprod.Recordset!costocaj, 0))
      DaoProdDbf.Recordset!peso = Val(IIf(Not IsNull(Adoprod.Recordset!peso), Adoprod.Recordset!peso, 0))
      DaoProdDbf.Recordset!barraspza = Val(IIf(Not IsNull(Adoprod.Recordset!barraspza), Adoprod.Recordset!barraspza, 0))
      DaoProdDbf.Recordset!barrascaja = Val(IIf(Not IsNull(Adoprod.Recordset!barrascaja), Adoprod.Recordset!barrascaja, 0))
      DaoProdDbf.Recordset!tasaieps = Val(IIf(Not IsNull(Adoprod.Recordset!tasaieps), Adoprod.Recordset!tasaieps, 0))
      DaoProdDbf.Recordset!claprove = IIf(Not IsNull(Adoprod.Recordset!claprove), Trim(Adoprod.Recordset!claprove), " ")
      DaoProdDbf.Recordset!proceden = Val(IIf(Not IsNull(Adoprod.Recordset!procedencia), Adoprod.Recordset!procedencia, 0))
      DaoProdDbf.Recordset!fecact = Adoprod.Recordset!fecact
      DaoProdDbf.Recordset!tantos = Val(IIf(Not IsNull(Adoprod.Recordset!cajas), Adoprod.Recordset!cajas, 0))
      DaoProdDbf.Recordset!entre = Val(IIf(Not IsNull(Adoprod.Recordset!encajas), Adoprod.Recordset!encajas, 0))
      DaoProdDbf.Recordset!descto01 = Val(IIf(Not IsNull(Adoprod.Recordset!decto1), Adoprod.Recordset!decto1, 0))
      DaoProdDbf.Recordset!descto02 = Val(IIf(Not IsNull(Adoprod.Recordset!decto2), Adoprod.Recordset!decto2, 0))
      DaoProdDbf.Recordset!descto03 = Val(IIf(Not IsNull(Adoprod.Recordset!decto3), Adoprod.Recordset!decto3, 0))
      DaoProdDbf.Recordset!descto04 = Val(IIf(Not IsNull(Adoprod.Recordset!decto4), Adoprod.Recordset!decto4, 0))
      DaoProdDbf.Recordset!deScto05 = Val(IIf(Not IsNull(Adoprod.Recordset!decto5), Adoprod.Recordset!decto5, 0))
      DaoProdDbf.Recordset!descto06 = Val(IIf(Not IsNull(Adoprod.Recordset!financiero), Adoprod.Recordset!financiero, 0))
      DaoProdDbf.Recordset!descefec = Val(IIf(Not IsNull(Adoprod.Recordset!efectivo), Adoprod.Recordset!efectivo, 0))
      DaoProdDbf.Recordset!porcargo = Val(IIf(Not IsNull(Adoprod.Recordset!cargo1), Adoprod.Recordset!cargo1, 0))
      DaoProdDbf.Recordset!otrosrec = Val(IIf(Not IsNull(Adoprod.Recordset!cargo5), Adoprod.Recordset!cargo5, 0))
      DaoProdDbf.Recordset!iva = Val(IIf(Not IsNull(Adoprod.Recordset!cargo3), Adoprod.Recordset!cargo3, 0))
      DaoProdDbf.Recordset!ieps = Val(IIf(Not IsNull(Adoprod.Recordset!cargo4), Adoprod.Recordset!cargo4, 0))
      DaoProdDbf.Recordset!fletes = Val(IIf(Not IsNull(Adoprod.Recordset!flete), Adoprod.Recordset!flete, 0))
      DaoProdDbf.Recordset!prepaque = Val(IIf(Not IsNull(Adoprod.Recordset!precio1), Adoprod.Recordset!precio1, 0))
      DaoProdDbf.Recordset!precaja = Val(IIf(Not IsNull(Adoprod.Recordset!precio2), Adoprod.Recordset!precio2, 0))
      DaoProdDbf.Recordset!prelib1 = Val(IIf(Not IsNull(Adoprod.Recordset!precio3), Adoprod.Recordset!precio3, 0))
      DaoProdDbf.Recordset!prelib2 = Val(IIf(Not IsNull(Adoprod.Recordset!precio4), Adoprod.Recordset!precio4, 0))
      DaoProdDbf.Recordset!gananpaq = Val(IIf(Not IsNull(Adoprod.Recordset!escala1), Adoprod.Recordset!escala1, 0))
      DaoProdDbf.Recordset!ganancaj = Val(IIf(Not IsNull(Adoprod.Recordset!escala2), Adoprod.Recordset!escala2, 0))
      DaoProdDbf.Recordset!gananlib1 = Val(IIf(Not IsNull(Adoprod.Recordset!escala3), Adoprod.Recordset!escala3, 0))
      DaoProdDbf.Recordset!gananlib2 = Val(IIf(Not IsNull(Adoprod.Recordset!escala4), Adoprod.Recordset!escala4, 0))
     ' DaoProdDbf.Recordset!CLAFAMIL = AdoProd.Recordset!CLAFAMIL
      DaoProdDbf.Recordset!oferta = Adoprod.Recordset!ofertado
      
      DaoProdDbf.Recordset.Update
      If respsn = vbNo Then
        cn.Execute "UPDATE tfproduc SET actualizado = 0 WHERE consec = '" & Trim(Adoprod.Recordset!consec) & "'"
      End If
      Adoprod.Recordset.MoveNext
    Wend


DaoProdDbf.Connect = "dbase III;"
DaoProdDbf.DatabaseName = "C:\paso"
DaoProdDbf.RecordsetType = Table
DaoProdDbf.RecordSource = "catprov"
DaoProdDbf.Refresh

Adoprov.Refresh
ProBar1.Min = 0
ProBar1.Max = Adoprov.Recordset.RecordCount
Lbltrans.Visible = True
Lbltrans.Refresh
ProBar1.Visible = True
v = 0
If Adoprov.Recordset.RecordCount > 0 Then
    Adoprov.Recordset.MoveFirst
    While Not Adoprov.Recordset.EOF
        v = v + 1
        ProBar1.Value = v
        DaoProdDbf.Recordset.AddNew
        DaoProdDbf.Recordset!PROVE = IIf(Not IsNull(Adoprov.Recordset!PROVE), Mid(Trim(Adoprov.Recordset!PROVE), 1, 3), " ")
        DaoProdDbf.Recordset!nomprove = IIf(Not IsNull(Adoprov.Recordset!nomprove), Trim(Adoprov.Recordset!nomprove), " ")
        DaoProdDbf.Recordset!dirpro = IIf(Not IsNull(Adoprov.Recordset!dirpro), Trim(Adoprov.Recordset!dirpro), " ")
        DaoProdDbf.Recordset!colpro = IIf(Not IsNull(Adoprov.Recordset!colpro), Trim(Adoprov.Recordset!colpro), " ")
        DaoProdDbf.Recordset!delpro = IIf(Not IsNull(Adoprov.Recordset!delpro), Trim(Adoprov.Recordset!delpro), " ")
        DaoProdDbf.Recordset!codpro = IIf(Not IsNull(Adoprov.Recordset!codpro), Trim(Adoprov.Recordset!codpro), " ")
        DaoProdDbf.Recordset!ciupro = IIf(Not IsNull(Adoprov.Recordset!ciupro), Trim(Adoprov.Recordset!ciupro), " ")
        DaoProdDbf.Recordset!locpro = IIf(Not IsNull(Adoprov.Recordset!locpro), Trim(Adoprov.Recordset!locpro), " ")
        DaoProdDbf.Recordset!telpro = IIf(Not IsNull(Adoprov.Recordset!telpro), Trim(Adoprov.Recordset!telpro), " ")
        DaoProdDbf.Recordset.Update
        Adoprov.Recordset.MoveNext
    Wend

End If


ProBar1.Visible = False
Lbltrans.Visible = False
Lbltrans.Refresh
DaoProdDbf.Recordset.Close


DaoProdDbf.Connect = "dbase III;"
DaoProdDbf.DatabaseName = "C:\paso"
DaoProdDbf.RecordsetType = Table
DaoProdDbf.RecordSource = "productox"
DaoProdDbf.Refresh
Else
    MsgBox "No existieron cambios", vbInformation
End If
Exit Sub
error:
MsgBox Err.Description
End Sub


Sub Importa()
On Error Resume Next

respsn = MsgBox("  Importación TOTAL", vbQuestion + vbYesNoCancel, "Utilerias")
If respsn = vbCancel Then Exit Sub


AdoTfproduc.CursorType = adOpenKeyset
AdoTfproduc.LockType = adLockOptimistic
AdoTfproduc.CommandType = adCmdText
AdoTfproduc.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
AdoTfproduc.RecordSource = "select * from tfproduc"
AdoTfproduc.Refresh

AdoCargos.CursorType = adOpenKeyset
AdoCargos.LockType = adLockOptimistic
AdoCargos.CommandType = adCmdText
AdoCargos.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
AdoCargos.RecordSource = "select * from CARGOS"
AdoCargos.Refresh

AdoDescuentos.CursorType = adOpenKeyset
AdoDescuentos.LockType = adLockOptimistic
AdoDescuentos.CommandType = adCmdText
AdoDescuentos.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
AdoDescuentos.RecordSource = "select * from DESCUENTOS"
AdoDescuentos.Refresh

Adopreprod.CursorType = adOpenKeyset
Adopreprod.LockType = adLockOptimistic
Adopreprod.CommandType = adCmdText
Adopreprod.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
Adopreprod.RecordSource = "select * from PREPROD"
Adopreprod.Refresh

AdoDescprod.CursorType = adOpenKeyset
AdoDescprod.LockType = adLockOptimistic
AdoDescprod.CommandType = adCmdText
AdoDescprod.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
AdoDescprod.RecordSource = "select * from descprod"
AdoDescprod.Refresh


AdoMargen.CursorType = adOpenKeyset
AdoMargen.LockType = adLockOptimistic
AdoMargen.CommandType = adCmdText
AdoMargen.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
AdoMargen.RecordSource = "select * from margen"
AdoMargen.Refresh

Dim v As Integer
Dim Clave As String

DaoProdDbf.Connect = "dbase III;"
DaoProdDbf.DatabaseName = "C:\paso"
DaoProdDbf.RecordsetType = Table
DaoProdDbf.RecordSource = "producto"
DaoProdDbf.Refresh

If DaoProdDbf.Recordset.RecordCount > 0 Then
    ProBar1.Min = 0
    ProBar1.Max = DaoProdDbf.Recordset.RecordCount
    Lbltrans.Visible = True
    Lbltrans.Refresh
    ProBar1.Visible = True
    v = 0
    DaoProdDbf.Recordset.MoveFirst
    AdoTfproduc.Refresh
    While Not DaoProdDbf.Recordset.EOF
      v = v + 1
      ProBar1.Value = v
            concomilla = InStr(1, DaoProdDbf.Recordset!descripc, "'")
            If concomilla > 0 Then
             xdescripc = Mid(DaoProdDbf.Recordset!descripc, 1, concomilla - 1) + "  " + Mid(DaoProdDbf.Recordset!descripc, concomilla + 1, Len(DaoProdDbf.Recordset!descripc))
            Else
             xdescripc = DaoProdDbf.Recordset!descripc
            End If
             xdescripc = Mid(xdescripc, 1, 50)
            
            concomilla = InStr(1, xdescripc, "'")
            If concomilla > 0 Then
             xdescripc2 = Mid(xdescripc, 1, concomilla - 1) + "  " + Mid(xdescripc, concomilla + 1, Len(xdescripc))
            Else
             xdescripc2 = xdescripc
            End If
             xdescripc = Mid(xdescripc2, 1, 50)

             
            concomillan = InStr(1, DaoProdDbf.Recordset!Nomcorto, "'")
            If concomillan > 0 Then
             xnomcorto = Mid(DaoProdDbf.Recordset!Nomcorto, 1, concomillan - 1) + "  " + Mid(DaoProdDbf.Recordset!Nomcorto, concomillan + 1, Len(DaoProdDbf.Recordset!Nomcorto))
            Else
             xnomcorto = DaoProdDbf.Recordset!Nomcorto
            End If
             xnomcorto = Mid(xnomcorto, 1, 20)
      Clave = Str(DaoProdDbf.Recordset!consec + 1000000)
      'PARA CUANDO SE EXPORTE DESDE LA MISMA BASE DE DATOS
      
      
      AdoTfproduc.Recordset.MoveFirst

      AdoTfproduc.Recordset.Find " CONSEC = " & Trim(Clave)
      If AdoTfproduc.Recordset.EOF Then  'Or AdoTfproduc.BOF Then
            cn.Execute "insert into tfproduc (CONSEC,DESCRIPC,NOMCORTO,CONTENID,FLETEX,FLESUB,MEDIDA,PAQUETES,COSTOPAQ,cOSTOCAJ,PESO,BARRASPZA,BARRASCAJA,tASAIEPS,CLAPROVE,procedencia,fecact,cajas,encajas) values " & _
            "('" & Trim(Clave) & "','" & IIf(Not IsNull(DaoProdDbf.Recordset!descripc), xdescripc, "A") & "','" & IIf(Not IsNull(DaoProdDbf.Recordset!Nomcorto), xnomcorto, "A") & "'," & IIf(Not IsNull(DaoProdDbf.Recordset!Contenid), DaoProdDbf.Recordset!Contenid, 0) & ", " & IIf(Not IsNull(DaoProdDbf.Recordset!fletex), DaoProdDbf.Recordset!fletex, 0) & _
            "," & IIf(Not IsNull(DaoProdDbf.Recordset!flesub), DaoProdDbf.Recordset!flesub, 0) & ",'" & IIf(Not IsNull(DaoProdDbf.Recordset!Medida), DaoProdDbf.Recordset!Medida, "A") & "'," & IIf(Not IsNull(DaoProdDbf.Recordset!Paquetes), DaoProdDbf.Recordset!Paquetes, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!costopaq), DaoProdDbf.Recordset!costopaq, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!prelista), DaoProdDbf.Recordset!prelista, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!peso), DaoProdDbf.Recordset!peso, 0) & "," & _
            IIf(Not IsNull(DaoProdDbf.Recordset!barraspza), DaoProdDbf.Recordset!barraspza, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!barrascaja), DaoProdDbf.Recordset!barrascaja, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!tasaieps), DaoProdDbf.Recordset!tasaieps, 0) & ",'" & IIf(Not IsNull(DaoProdDbf.Recordset!claprove), DaoProdDbf.Recordset!claprove, 0) & "',1,'" & DaoProdDbf.Recordset!fecact & "'," & IIf(Not IsNull(DaoProdDbf.Recordset!tantos), DaoProdDbf.Recordset!tantos, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!entre), DaoProdDbf.Recordset!entre, 0) & ")"
        Else
            If respsn = vbNo Then
                cn.Execute " UPDATE TFPRODUC SET FLETEX = " & IIf(Not IsNull(DaoProdDbf.Recordset!fletex), DaoProdDbf.Recordset!fletex, 0) & " , FLESUB = " & IIf(Not IsNull(DaoProdDbf.Recordset!flesub), DaoProdDbf.Recordset!flesub, 0) & _
                ", COSTOPAQ =" & IIf(Not IsNull(DaoProdDbf.Recordset!costopaq), DaoProdDbf.Recordset!costopaq, 0) & ", COSTOCAJ =" & IIf(Not IsNull(DaoProdDbf.Recordset!prelista), DaoProdDbf.Recordset!prelista, 0) & ", PESO = " & IIf(Not IsNull(DaoProdDbf.Recordset!peso), DaoProdDbf.Recordset!peso, 0) & "," & _
                "BARRASPZA = " & IIf(Not IsNull(DaoProdDbf.Recordset!barraspza), DaoProdDbf.Recordset!barraspza, 0) & ", BARRASCAJA =" & IIf(Not IsNull(DaoProdDbf.Recordset!barrascaja), DaoProdDbf.Recordset!barrascaja, 0) & ",TASAIEPS = " & IIf(Not IsNull(DaoProdDbf.Recordset!tasaieps), DaoProdDbf.Recordset!tasaieps, 0) & ",CLAPROVE = '" & IIf(Not IsNull(DaoProdDbf.Recordset!claprove), DaoProdDbf.Recordset!claprove, 0) & "'," & _
                "cajas = " & IIf(Not IsNull(DaoProdDbf.Recordset!tantos), DaoProdDbf.Recordset!tantos, 0) & ", encajas= " & IIf(Not IsNull(DaoProdDbf.Recordset!entre), DaoProdDbf.Recordset!entre, 0) & "," & _
                "ACTIVO = 1 Where CONSEC = '" & Trim(Clave) & "'"
            Else
                cn.Execute " UPDATE TFPRODUC SET DESCRIPC = '" & IIf(Not IsNull(DaoProdDbf.Recordset!descripc), xdescripc, "A") & "', NOMCORTO = '" & IIf(Not IsNull(DaoProdDbf.Recordset!Nomcorto), xnomcorto, "A") & "', CONTENID = " & IIf(Not IsNull(DaoProdDbf.Recordset!Contenid), DaoProdDbf.Recordset!Contenid, 0) & "," & _
                "FLETEX = " & IIf(Not IsNull(DaoProdDbf.Recordset!fletex), DaoProdDbf.Recordset!fletex, 0) & " , FLESUB = " & IIf(Not IsNull(DaoProdDbf.Recordset!flesub), DaoProdDbf.Recordset!flesub, 0) & ", MEDIDA = '" & IIf(Not IsNull(DaoProdDbf.Recordset!Medida), DaoProdDbf.Recordset!Medida, 0) & "', PAQUETES = " & IIf(Not IsNull(DaoProdDbf.Recordset!Paquetes), DaoProdDbf.Recordset!Paquetes, 0) & "," & _
                " COSTOPAQ =" & IIf(Not IsNull(DaoProdDbf.Recordset!costopaq), DaoProdDbf.Recordset!costopaq, 0) & ", COSTOCAJ =" & IIf(Not IsNull(DaoProdDbf.Recordset!prelista), DaoProdDbf.Recordset!prelista, 0) & ", PESO = " & IIf(Not IsNull(DaoProdDbf.Recordset!peso), DaoProdDbf.Recordset!peso, 0) & "," & _
                "BARRASPZA = " & IIf(Not IsNull(DaoProdDbf.Recordset!barraspza), DaoProdDbf.Recordset!barraspza, 0) & ", BARRASCAJA =" & IIf(Not IsNull(DaoProdDbf.Recordset!barrascaja), DaoProdDbf.Recordset!barrascaja, 0) & ",TASAIEPS = " & IIf(Not IsNull(DaoProdDbf.Recordset!tasaieps), DaoProdDbf.Recordset!tasaieps, 0) & ",CLAPROVE = '" & IIf(Not IsNull(DaoProdDbf.Recordset!claprove), DaoProdDbf.Recordset!claprove, 0) & "'," & _
                "cajas = " & IIf(Not IsNull(DaoProdDbf.Recordset!tantos), DaoProdDbf.Recordset!tantos, 0) & ", encajas= " & IIf(Not IsNull(DaoProdDbf.Recordset!entre), DaoProdDbf.Recordset!entre, 0) & "," & _
                "ACTIVO = 1 Where CONSEC = '" & Trim(Clave) & "'"
                End If
        End If
      
      
      
      AdoCargos.Recordset.MoveFirst
      AdoCargos.Recordset.Find " caprod = " & Trim(Clave)
        
      If AdoCargos.Recordset.EOF Then
 
        cn.Execute "insert into cargos (CAPROD,cargo1,cargo_efectivo,iva,ieps,flete_efectivo) values ('" & Trim(Clave) & "'," & IIf(Not IsNull(DaoProdDbf.Recordset!porcargo), DaoProdDbf.Recordset!porcargo, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!otrosrec), DaoProdDbf.Recordset!otrosrec, 0) & _
        "," & IIf(Not IsNull(DaoProdDbf.Recordset!iva), DaoProdDbf.Recordset!iva, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!ieps), DaoProdDbf.Recordset!ieps, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!fletes), DaoProdDbf.Recordset!fletes, 0) & ")"
      Else
        cn.Execute "update cargos set cargo1 = " & IIf(Not IsNull(DaoProdDbf.Recordset!porcargo), DaoProdDbf.Recordset!porcargo, 0) & ", cargo_efectivo = " & IIf(Not IsNull(DaoProdDbf.Recordset!otrosrec), DaoProdDbf.Recordset!otrosrec, 0) & _
        ", iva = " & IIf(Not IsNull(DaoProdDbf.Recordset!iva), DaoProdDbf.Recordset!iva, 0) & ", ieps = " & IIf(Not IsNull(DaoProdDbf.Recordset!ieps), DaoProdDbf.Recordset!ieps, 0) & ", flete_efectivo = " & IIf(Not IsNull(DaoProdDbf.Recordset!fletes), DaoProdDbf.Recordset!fletes, 0) & _
        "Where CAPROD = '" & Trim(Clave) & "'"
      End If
      Adopreprod.Recordset.MoveFirst
      Adopreprod.Recordset.Find " preclave = " & Trim(Clave)
      
        
      If Adopreprod.Recordset.EOF Then
        cn.Execute "insert into preprod (PRECLAVE,precio1,precio2,precio3,precio4,fechaact) values ( '" & Trim(Clave) & "'," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!prepaque), DaoProdDbf.Recordset!prepaque, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!precaja), DaoProdDbf.Recordset!precaja, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!prelib1), DaoProdDbf.Recordset!prelib1, 0) & _
        "," & IIf(Not IsNull(DaoProdDbf.Recordset!prelib2), DaoProdDbf.Recordset!prelib2, 0) & ", '" & IIf(Not IsNull(DaoProdDbf.Recordset!fecact), DaoProdDbf.Recordset!fecact, Date) & "')"
      Else
        cn.Execute "update preprod set precio1 = " & IIf(Not IsNull(DaoProdDbf.Recordset!prepaque), DaoProdDbf.Recordset!prepaque, 0) & ",precio2 = " & IIf(Not IsNull(DaoProdDbf.Recordset!precaja), DaoProdDbf.Recordset!precaja, 0) & _
        ",precio3 = " & IIf(Not IsNull(DaoProdDbf.Recordset!prelib1), DaoProdDbf.Recordset!prelib1, 0) & ", precio4 = " & IIf(Not IsNull(DaoProdDbf.Recordset!prelib2), DaoProdDbf.Recordset!prelib2, 0) & _
        ",fechaact = '" & IIf(Not IsNull(DaoProdDbf.Recordset!fecact), DaoProdDbf.Recordset!fecact, Date) & _
        "' Where PRECLAVE = '" & Trim(Clave) & "'"
      End If
      
      AdoDescuentos.Recordset.MoveFirst
      AdoDescuentos.Recordset.Find " deprod = " & Trim(Clave)
      
        
      If AdoDescuentos.Recordset.EOF Then
        cn.Execute "insert into descuentos (DEPROD,decto1,decto2,decto3,dectooferta,dectofinanciero,dectoefectivo,decto5) values ('" & Trim(Clave) & "'," & IIf(Not IsNull(DaoProdDbf.Recordset!descto01), DaoProdDbf.Recordset!descto01, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!descto02), DaoProdDbf.Recordset!descto02, 0) & "," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!descto03), DaoProdDbf.Recordset!descto03, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!descto04), DaoProdDbf.Recordset!descto04, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!descto06), DaoProdDbf.Recordset!descto06, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!descefec), DaoProdDbf.Recordset!descefec, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!deScto05), DaoProdDbf.Recordset!deScto05, 0) & ")"
      Else
        cn.Execute "update descuentos set decto1 = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto01), DaoProdDbf.Recordset!descto01, 0) & ", decto2 = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto02), DaoProdDbf.Recordset!descto02, 0) & ", decto3 = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto03), DaoProdDbf.Recordset!descto03, 0) & _
        ", dectooferta = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto04), DaoProdDbf.Recordset!descto04, 0) & ", dectofinanciero = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto06), DaoProdDbf.Recordset!descto06, 0) & ", dectoefectivo = " & IIf(Not IsNull(DaoProdDbf.Recordset!descefec), DaoProdDbf.Recordset!descefec, 0) & _
        ", decto5= " & IIf(Not IsNull(DaoProdDbf.Recordset!deScto05), DaoProdDbf.Recordset!deScto05, 0) & " Where DEPROD = '" & Trim(Clave) & "'"
      End If
      AdoDescprod.Recordset.MoveFirst
      AdoDescprod.Recordset.Find " producto = " & Trim(Clave)
      
      nprecio = CalCosto()
      If AdoDescprod.Recordset.EOF Then
        cn.Execute "insert into descprod (proveedor,producto,decto1,decto2,decto3,decto4," & _
        "financiero,efectivo,cargo1,cargo5,cargo3,cargo4,FLETE,cajas,encajas,decto5,preciolista,costo) " & _
        " values ('" & DaoProdDbf.Recordset!claprove & "','" & Trim(Clave) & "'," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!descto01), DaoProdDbf.Recordset!descto01, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!descto02), DaoProdDbf.Recordset!descto02, 0) & "," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!descto03), DaoProdDbf.Recordset!descto03, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!descto04), DaoProdDbf.Recordset!descto04, 0) & "," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!descto06), DaoProdDbf.Recordset!descto06, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!descefec), DaoProdDbf.Recordset!descefec, 0) & "," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!porcargo), DaoProdDbf.Recordset!porcargo, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!otrosrec), DaoProdDbf.Recordset!otrosrec, 0) & "," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!iva), DaoProdDbf.Recordset!iva, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!ieps), DaoProdDbf.Recordset!ieps, 0) & "," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!fletes), DaoProdDbf.Recordset!fletes, 0) & "," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!tantos), DaoProdDbf.Recordset!tantos, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!entre), DaoProdDbf.Recordset!entre, 0) & "," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!deScto05), DaoProdDbf.Recordset!deScto05, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!prelista), DaoProdDbf.Recordset!prelista, 0) & _
        "," & nprecio & ")"
      Else
        cn.Execute "update descprod set decto1 = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto01), DaoProdDbf.Recordset!descto01, 0) & ", decto2 = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto02), DaoProdDbf.Recordset!descto02, 0) & ", decto3 = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto03), DaoProdDbf.Recordset!descto03, 0) & _
        ", decto4 = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto04), DaoProdDbf.Recordset!descto04, 0) & ", financiero = " & IIf(Not IsNull(DaoProdDbf.Recordset!descto06), DaoProdDbf.Recordset!descto06, 0) & ", efectivo = " & IIf(Not IsNull(DaoProdDbf.Recordset!descefec), DaoProdDbf.Recordset!descefec, 0) & _
        ", cargo1 = " & IIf(Not IsNull(DaoProdDbf.Recordset!porcargo), DaoProdDbf.Recordset!porcargo, 0) & ", cargo5 = " & IIf(Not IsNull(DaoProdDbf.Recordset!otrosrec), DaoProdDbf.Recordset!otrosrec, 0) & ",proveedor = '" & DaoProdDbf.Recordset!claprove & "'" & _
        ", cargo3 = " & IIf(Not IsNull(DaoProdDbf.Recordset!iva), DaoProdDbf.Recordset!iva, 0) & ", cargo4 = " & IIf(Not IsNull(DaoProdDbf.Recordset!ieps), DaoProdDbf.Recordset!ieps, 0) & ", FLETE = " & IIf(Not IsNull(DaoProdDbf.Recordset!fletes), DaoProdDbf.Recordset!fletes, 0) & _
       ", decto5 = " & IIf(Not IsNull(DaoProdDbf.Recordset!deScto05), DaoProdDbf.Recordset!deScto05, 0) & ",preCIOlista=" & IIf(Not IsNull(DaoProdDbf.Recordset!prelista), DaoProdDbf.Recordset!prelista, 0) & " Where producto = '" & Trim(Clave) & "'"
      End If
      
      
      AdoMargen.Recordset.MoveFirst
      AdoMargen.Recordset.Find " producto = " & Trim(Clave)
      
        
      If AdoMargen.Recordset.EOF Then
        cn.Execute "insert into margen (producto,escala1,escala2,escala3,escala4) values ( '" & Trim(Clave) & "'," & _
        IIf(Not IsNull(DaoProdDbf.Recordset!gananpaq), DaoProdDbf.Recordset!gananpaq, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!ganancaj), DaoProdDbf.Recordset!ganancaj, 0) & "," & IIf(Not IsNull(DaoProdDbf.Recordset!gananlib1), DaoProdDbf.Recordset!gananlib1, 0) & _
        "," & IIf(Not IsNull(DaoProdDbf.Recordset!gananlib2), DaoProdDbf.Recordset!gananlib2, 0) & ")"
      Else
        cn.Execute "update margen set escala1 = " & IIf(Not IsNull(DaoProdDbf.Recordset!gananpaq), DaoProdDbf.Recordset!gananpaq, 0) & ",escala2 = " & IIf(Not IsNull(DaoProdDbf.Recordset!ganancaj), DaoProdDbf.Recordset!ganancaj, 0) & _
        ",escala3 = " & IIf(Not IsNull(DaoProdDbf.Recordset!gananlib1), DaoProdDbf.Recordset!gananlib1, 0) & ", escala4 = " & IIf(Not IsNull(DaoProdDbf.Recordset!gananlib2), DaoProdDbf.Recordset!gananlib2, 0) & _
        " Where producto = '" & Trim(Clave) & "'"
      End If
      
      DaoProdDbf.Recordset.MoveNext
      
    Wend
End If

DaoProdDbf.Connect = "dbase III;"
DaoProdDbf.DatabaseName = "C:\paso"
DaoProdDbf.RecordsetType = Table
DaoProdDbf.RecordSource = "catprov"
DaoProdDbf.Refresh
 
 Clave = ""
If DaoProdDbf.Recordset.RecordCount > 0 Then
    ProBar1.Min = 0
    ProBar1.Max = DaoProdDbf.Recordset.RecordCount
    Lbltrans.Visible = True
    Lbltrans.Refresh
    ProBar1.Visible = True
    v = 0
    DaoProdDbf.Recordset.MoveFirst
    Adoprov.Refresh
    While Not DaoProdDbf.Recordset.EOF
      v = v + 1
      ProBar1.Value = v
      Clave = DaoProdDbf.Recordset!PROVE
      concomilla = InStr(1, DaoProdDbf.Recordset!nomprove, "'")
      If concomilla > 0 Then
         xnomprove = Mid(DaoProdDbf.Recordset!nomprove, 1, concomilla - 1) + "  " + Mid(DaoProdDbf.Recordset!nomprove, concomilla + 1, Len(DaoProdDbf.Recordset!nomprove))
      Else
         xnomprove = DaoProdDbf.Recordset!nomprove
      End If
            
      Adoprov.Recordset.MoveFirst
      Adoprov.Recordset.Find "prove = '" & Trim(Clave) & "'"
      If Adoprov.Recordset.EOF Then
        cn.Execute "INSERT INTO CATPROV (prove,nomprove,dirpro,colpro,delpro,codpro,ciupro,locpro,telpro) values ('" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!PROVE), DaoProdDbf.Recordset!PROVE, 0) & "','" & IIf(Not IsNull(DaoProdDbf.Recordset!nomprove), xnomprove, 0) & "','" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!dirpro), DaoProdDbf.Recordset!dirpro, 0) & "','" & IIf(Not IsNull(DaoProdDbf.Recordset!colpro), DaoProdDbf.Recordset!colpro, 0) & "','" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!delpro), DaoProdDbf.Recordset!delpro, 0) & "','" & IIf(Not IsNull(DaoProdDbf.Recordset!codpro), DaoProdDbf.Recordset!codpro, 0) & "','" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!ciupro), DaoProdDbf.Recordset!ciupro, 0) & "','" & IIf(Not IsNull(DaoProdDbf.Recordset!locpro), DaoProdDbf.Recordset!locpro, 0) & "','" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!telpro), DaoProdDbf.Recordset!telpro, 0) & "')"
      Else
        cn.Execute "UPDATE catprov SET nomprove = '" & IIf(Not IsNull(DaoProdDbf.Recordset!nomprove), xnomprove, 0) & "', dirpro = '" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!dirpro), DaoProdDbf.Recordset!dirpro, 0) & "', colpro = '" & IIf(Not IsNull(DaoProdDbf.Recordset!colpro), DaoProdDbf.Recordset!colpro, 0) & "', delpro =  '" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!delpro), DaoProdDbf.Recordset!delpro, 0) & "', codpro =  '" & IIf(Not IsNull(DaoProdDbf.Recordset!codpro), DaoProdDbf.Recordset!codpro, 0) & "', ciupro = '" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!ciupro), DaoProdDbf.Recordset!ciupro, 0) & "', locpro = '" & IIf(Not IsNull(DaoProdDbf.Recordset!locpro), DaoProdDbf.Recordset!locpro, 0) & "', telpro = '" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!telpro), DaoProdDbf.Recordset!telpro, 0) & "' where prove = '" & Trim(Clave) & "'"
            
      End If
      DaoProdDbf.Recordset.MoveNext
      
     Wend
End If

ProBar1.Visible = False
Lbltrans.Visible = False
Lbltrans.Refresh

DaoProdDbf.Connect = "dbase III;"
DaoProdDbf.DatabaseName = "C:\paso"
DaoProdDbf.RecordsetType = Table
DaoProdDbf.RecordSource = "productox"
DaoProdDbf.Refresh

Exit Sub
error:
MsgBox Err.Description
End Sub





Function CalCosto() As Double
Dim nprecio As Double
    nprecio = 0
    nprecio = IIf(Not IsNull(DaoProdDbf.Recordset!prelista), DaoProdDbf.Recordset!prelista, 0)
     
    'calcula cargos %
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!iva), DaoProdDbf.Recordset!iva, 0)
    If npreciopaso > 0 Then
    nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!ieps), DaoProdDbf.Recordset!ieps, 0)
    If npreciopaso > 0 Then
    nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
      
    nprecio = Round(nprecio, 2)
        
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto01), DaoProdDbf.Recordset!descto01, 0)
    If npreciopaso > 0 Then
    nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto02), DaoProdDbf.Recordset!descto02, 0)
    If npreciopaso > 0 Then
    nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto03), DaoProdDbf.Recordset!descto03, 0)
    If npreciopaso > 0 Then
    nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto04), DaoProdDbf.Recordset!descto04, 0)
    If npreciopaso > 0 Then
    nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
    
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto06), DaoProdDbf.Recordset!descto06, 0)
    If npreciopaso > 0 Then
    nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
    
    'descuento efectivo
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descefec), DaoProdDbf.Recordset!descefec, 0)
    If npreciopaso > 0 Then
    nprecio = nprecio - npreciopaso
    End If
    
    nprecio = Round(nprecio, 2)

    CalCosto = nprecio
    
End Function
