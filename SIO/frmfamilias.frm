VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ffamilia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogo de Familias"
   ClientHeight    =   5145
   ClientLeft      =   1350
   ClientTop       =   1530
   ClientWidth     =   10770
   Icon            =   "frmfamilias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   41
      Top             =   4815
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "                                                                                Para salir presione la tecla       [ Esc ]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adolinasig 
      Height          =   330
      Left            =   0
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
   Begin MSAdodcLib.Adodc Adofamilia 
      Height          =   375
      Left            =   480
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   10455
      Begin VB.CommandButton Cmdafam 
         Caption         =   "&Catalogo de Lineas"
         Height          =   495
         Index           =   2
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Cmdver 
         Caption         =   "&Lineas en Familia"
         Height          =   495
         Index           =   0
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Cmdbus 
         Height          =   495
         Left            =   2640
         Picture         =   "frmfamilias.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Buscar Registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   495
         Left            =   6480
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdmodifica 
         Caption         =   "&Modificar"
         Height          =   495
         Left            =   5520
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdsalir 
         Height          =   495
         Left            =   9720
         Picture         =   "frmfamilias.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Salir del modulo"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdcancela 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   8520
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Height          =   495
         Left            =   240
         Picture         =   "frmfamilias.frx":06AE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ir al primer registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   840
         Picture         =   "frmfamilias.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ir al registro anterior"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   1440
         Picture         =   "frmfamilias.frx":0992
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ir al siguiente registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Height          =   495
         Left            =   2040
         Picture         =   "frmfamilias.frx":0B04
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ir al ultimo registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7560
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin MSAdodcLib.Adodc adodeptos 
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
         RecordSource    =   "select * from departamento"
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
      Begin MSAdodcLib.Adodc adolineas 
         Height          =   375
         Left            =   240
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   "select * from familias order by fdescrip"
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lineas asignadas "
      Height          =   4215
      Left            =   600
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   9375
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmfamilias.frx":0C76
         Height          =   3615
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "sfclave"
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
            DataField       =   "sfdescrip"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               DividerStyle    =   0
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   5595.024
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Cmdver 
         Caption         =   "Re&porte"
         Height          =   495
         Index           =   3
         Left            =   7800
         Picture         =   "frmfamilias.frx":0C8F
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Cmdver 
         Caption         =   "&Regresar"
         Height          =   495
         Index           =   2
         Left            =   7800
         Picture         =   "frmfamilias.frx":11C1
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame frmbusca 
      Height          =   4575
      Left            =   960
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   8655
      Begin VB.ListBox Lstlin 
         Height          =   3375
         Left            =   600
         TabIndex        =   27
         Top             =   600
         Width           =   7575
      End
      Begin VB.CommandButton Cmdver 
         Caption         =   "&Regresar"
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   26
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Familias Registradas en el Catalogo"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.TextBox txtobserva 
         Height          =   1575
         Left            =   2520
         TabIndex        =   3
         Top             =   1200
         Width           =   6975
      End
      Begin VB.TextBox txtdesc 
         Height          =   375
         Left            =   2520
         MaxLength       =   80
         TabIndex        =   2
         Top             =   720
         Width           =   6975
      End
      Begin VB.TextBox txtclave 
         Height          =   405
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   3000
         Width           =   5535
      End
      Begin VB.TextBox txtfdepto 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   720
         Picture         =   "frmfamilias.frx":1333
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ir al primer registro"
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones:"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Clave:"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Departamento:"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Asignar Lineas"
      Height          =   4575
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.ListBox Lstfamilia 
         Height          =   450
         ItemData        =   "frmfamilias.frx":14A5
         Left            =   7560
         List            =   "frmfamilias.frx":14A7
         TabIndex        =   31
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmfamilias.frx":14A9
         Height          =   3190
         Left            =   480
         TabIndex        =   34
         Top             =   600
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5636
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1.5
         RowHeight       =   15
         TabAction       =   1
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "sfclave"
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
            DataField       =   "sfdescrip"
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
            DataField       =   "sffamilia"
            Caption         =   "Familia"
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
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   5504.882
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Cmdafam 
         Caption         =   "&Cambiar Familia"
         Height          =   500
         Index           =   3
         Left            =   5880
         Picture         =   "frmfamilias.frx":14C2
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton Cmdafam 
         Caption         =   "&Buscar"
         Height          =   500
         Index           =   0
         Left            =   7440
         Picture         =   "frmfamilias.frx":1B2C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Cmdafam 
         Caption         =   "&Regresar"
         Height          =   500
         Index           =   1
         Left            =   8760
         Picture         =   "frmfamilias.frx":1C9E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "PARA ASIGNAR UNA FAMILIA A ESTA LINEA MODIFICAR LA CLAVE DE LA COLUMNA   "" Familia """
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
         Height          =   495
         Left            =   480
         TabIndex        =   39
         Top             =   3960
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "PERTENECE A LA :"
         Height          =   255
         Left            =   7320
         TabIndex        =   36
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "CATALOGO DE LINEAS"
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   360
         Width           =   3975
      End
   End
End
Attribute VB_Name = "ffamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lnuevo As Boolean


Private Sub Cmdbus_Click()
On Error GoTo Error:
Frame1.Visible = False
Frame2.Visible = False
frmbusca.Visible = True
Lstlin.Clear
    adolineas.Refresh
    Do While Not adolineas.Recordset.EOF = True
               If Not IsNull(adolineas.Recordset!fdescrip) Then
                 Lstlin.AddItem adolineas.Recordset!fdescrip + "  [" + adolineas.Recordset!fclave + "]"
               End If
              adolineas.Recordset.MoveNext
            Loop
           adolineas.Recordset.MoveFirst

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdcancela_Click()
On Error GoTo Error:
    adolineas.Refresh
    Call asigna
    Call habilitar
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdmodifica_Click()
On Error GoTo Error:
lnuevo = False
txtclave.Locked = True
txtdesc.SetFocus
Call dhabilitar
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub cmdnuevo_Click()
On Error GoTo Error:
 txtdesc.Text = ""
 txtobserva.Text = ""
 txtfdepto.Text = ""
 Combo2.Text = ""
 lnuevo = True
  Call nuevalin
  Call dhabilitar
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

Private Sub Cmdver_Click(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0
    frmbusca.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Frame4.Visible = False
    Frame3.Visible = True
Case 1
    frmbusca.Visible = False
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
Case 2
    frmbusca.Visible = False
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
Case 3
    fmenu.CR1.WindowTitle = "Reporte de Familia " & Trim(txtdesc.Text)
    fmenu.CR1.ReportFileName = App.Path & "\depfamlin.rpt"
    fmenu.CR1.Formulas(1) = "FORMSELEC = {FAMILIAS.fclave} = '" & Trim(txtclave.Text) & "'"
    fmenu.CR1.Connect = strconnect
    fmenu.CR1.WindowState = crptMaximized
    fmenu.CR1.Action = 1
End Select
Exit Sub
Error:
    MsgBox Err.Description

End Sub

Private Sub Combo2_Click()
If Trim(Combo2.Text) <> "" Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Command6_Click()
On Error GoTo Error:
If adolineas.Recordset.EOF = False And adolineas.Recordset.BOF = False Then
adolineas.Recordset.MoveFirst
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command2_Click()
On Error GoTo Error:
Dim reg As Integer
reg = adolineas.Recordset.AbsolutePosition
If reg > 1 Then
adolineas.Recordset.MovePrevious
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

reg = adolineas.Recordset.AbsolutePosition
treg = adolineas.Recordset.RecordCount
If reg < treg Then

adolineas.Recordset.MoveNext
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command4_Click()
On Error GoTo Error:
If adolineas.Recordset.EOF = False And adolineas.Recordset.BOF = False Then
adolineas.Recordset.MoveLast
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command5_Click()
On Error GoTo Error:
If Trim(txtdesc.Text) <> "" And Trim(txtfdepto.Text) <> "" Then
If lnuevo Then
    adolineas.Recordset.AddNew
End If
adolineas.Recordset!fclave = txtclave.Text
adolineas.Recordset!fdescrip = txtdesc.Text
adolineas.Recordset!fobserva = txtobserva.Text
adolineas.Recordset!fdepto = txtfdepto.Text
adolineas.Recordset.Update

Call habilitar
Else
    MsgBox "Favor de completar los datos ..."
    txtdesc.SetFocus
    
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
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
Private Sub DataGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
On Error GoTo Error:
 
 If ColIndex = 2 Then
       Cancel = True
        DataGrid2_ButtonClick (ColIndex)
   End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub DataGrid2_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error:
Dim L As ListBox
    Select Case ColIndex
       Case 2
            Set L = Lstfamilia
    If ColIndex = -1 Then Exit Sub
      With L
          'Abajo (3):
          .Left = DataGrid2.Left + DataGrid2.Columns(ColIndex).Left
          .Top = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + DataGrid2.RowHeight
          '.Width = dbgrdDetPed.Columns(ColIndex).Width + 15
          .ListIndex = 0
          .Visible = True
          .ZOrder 0
          .SetFocus
    End With
   End Select
   Exit Sub
Error:
   MsgBox Err.Description


End Sub



Private Sub asigna()
On Error GoTo Error:
If adolineas.Recordset.EOF = False And adolineas.Recordset.BOF = False Then

txtclave.Text = adolineas.Recordset!fclave
txtdesc.Text = adolineas.Recordset!fdescrip
txtobserva.Text = adolineas.Recordset!fobserva
txtfdepto.Text = adolineas.Recordset!fdepto
adofamilia.CommandType = adCmdText
adofamilia.CursorType = adOpenKeyset
adofamilia.LockType = adLockOptimistic
adofamilia.ConnectionString = cn.ConnectionString
adofamilia.RecordSource = "select * from lineas where sffamilia = '" & Trim(txtclave.Text) & "'"
adofamilia.Refresh

Lstfamilia.Clear
Lstfamilia.AddItem "                 "
Lstfamilia.AddItem Trim(txtdesc.Text)




txtfdepto_KeyPress 13
lnuevo = False

End If

'If Trim(txtclave.Text) <> "" Then
'    Lblfamiliax.Caption = txtdesc.Text'
'Else
'    Lblfamiliax.Caption = ""
'End If
'Lblfamiliax.Refresh


Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub nuevalin()
Dim rs As ADODB.Recordset
On Error GoTo Error:
If adolineas.Recordset.EOF = False And adolineas.Recordset.BOF = False Then
    'adolineas.Recordset.MoveLast
    'txtclave.Text = Right("0000" + Trim(Str(Val(adolineas.Recordset!fclave) + 1)), 3)
    Set rs = New ADODB.Recordset
    rs.Open "SELECT max(fclave) AS NvaCve from familias", cn, adOpenDynamic, adLockOptimistic, adCmdText
    txtclave.Text = Right("0000" + Trim(Str(rs!NvaCve + 1)), 3)
    txtclave.Locked = True
    txtdesc.SetFocus
Else
    txtclave.Text = "001"
    txtclave.Locked = True
    txtdesc.SetFocus
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
 Command6.Enabled = False
 Command2.Enabled = False
 Command3.Enabled = False
 Command4.Enabled = False
 Cmdbus.Enabled = False
 Cmdver(0).Enabled = False
Exit Sub
Error:
MsgBox Err.Description
 
End Sub
Private Sub habilitar()
On Error GoTo Error:
 Command5.Enabled = False
 cmdcancela.Enabled = False
 cmdnuevo.Enabled = True
 cmdmodifica.Enabled = True
 Command6.Enabled = True
 Command2.Enabled = True
 Command3.Enabled = True
 Command4.Enabled = True
 Cmdbus.Enabled = True
 Cmdver(0).Enabled = True
Exit Sub
Error:
MsgBox Err.Description
 End Sub



Private Sub cmdgraba_Click()
On Error GoTo Error:
adolineas.Recordset!fdepto = txtfdepto.Text
adolineas.Recordset.Update
adolineas.Refresh
Exit Sub
Error:
MsgBox Err.Description
End Sub



Private Sub Lstlin_DblClick()
On Error GoTo Error:
Lstlin_KeyPress 13
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Lstlin_KeyPress(KeyAscii As Integer)
On Error GoTo Error:

If KeyAscii = 13 Then
Dim N As Integer
Dim Cvelin As String

             N = InStr(1, Lstlin.List(Lstlin.ListIndex), "[")
             Cvelin = Mid(Lstlin.List(Lstlin.ListIndex), N + 1, Len(Lstlin.List(Lstlin.ListIndex)) - N - 1)
             adolineas.Recordset.MoveFirst
             adolineas.Recordset.Find "fclave = '" & Cvelin & "'"
             If adolineas.Recordset.EOF Then
                MsgBox "   "
                            
             End If
             

frmbusca.Visible = False
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = False
Call asigna

End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtclave_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
            If Trim(txtclave.Text) <> "" Then
          
             adolineas.Recordset.MoveFirst
             adolineas.Recordset.Find "fclave = '" & Trim(txtclave.Text) & "'"
             Call asigna
             End If
    txtdesc.SetFocus
    End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
txtobserva.SetFocus
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub combo2_lostFocus()
On Error GoTo Error:
Dim Cvedepto As String
Dim N As Integer
            If Trim(Combo2.Text) <> "" Then
            If Combo2.ListIndex = -1 Then Exit Sub
             N = InStr(1, Combo2.List(Combo2.ListIndex), "[")
             Cvedepto = Mid(Combo2.List(Combo2.ListIndex), N + 1, Len(Combo2.List(Combo2.ListIndex)) - N - 1)
             adodeptos.Recordset.MoveFirst
             adodeptos.Recordset.Find "depclave = '" & Cvedepto & "'"
             If adodeptos.Recordset.EOF Then
                MsgBox "Seleccione un departamento "
                Combo2.SetFocus
             Else
                         
             txtfdepto.Text = adodeptos.Recordset!depclave
             End If
             Else
             MsgBox "Seleccione un departamento "
                Combo2.SetFocus
             End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Error:
adodeptos.CursorType = adOpenKeyset
adodeptos.CommandType = adCmdText
adodeptos.RecordSource = "select * from departamento"
adodeptos.ConnectionString = strconnect

Combo2.Clear
    adodeptos.Refresh
    Do While Not adodeptos.Recordset.EOF = True
               If Not IsNull(adodeptos.Recordset!depdescrip) Then
                 Combo2.AddItem adodeptos.Recordset!depdescrip + "  [" + adodeptos.Recordset!depclave + "]"
               End If
              adodeptos.Recordset.MoveNext
            Loop
           adodeptos.Recordset.MoveFirst

adolineas.ConnectionString = strconnect
adolineas.Refresh


Adolinasig.CursorType = adOpenKeyset
Adolinasig.CommandType = adCmdText
Adolinasig.RecordSource = "select * from lineas order by sfdescrip"
Adolinasig.ConnectionString = strconnect
Adolinasig.Refresh

Call asigna
Exit Sub
Error:
MsgBox Err.Description
           
End Sub

Private Sub txtfdepto_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
            If Trim(txtfdepto.Text) <> "" Then
             adodeptos.Refresh
             adodeptos.Recordset.MoveFirst
             adodeptos.Recordset.Find "depclave = '" & Trim(txtfdepto.Text) & "'"
             If adodeptos.Recordset.EOF Then
                MsgBox "Seleccione un departamento "
                Combo2.SetFocus
             Else
                Combo2.Text = adodeptos.Recordset!depdescrip + "[" + adodeptos.Recordset!depclave + "]"
                txtfdepto.Text = adodeptos.Recordset!depclave
                Combo2.Refresh
             End If
             Else
             MsgBox "Seleccione un departamento "
                Combo2.SetFocus
             End If
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub
Private Sub Lstfamilia_DblClick()
Dim Antes
Antes = DataGrid2.Bookmark
If Trim(Lstfamilia.Text) = "" Then
    Adolinasig.Refresh
    DataGrid2.Bookmark = Antes
    
Else
    DataGrid2.Columns(2).Text = Trim(txtclave.Text)
End If
Lstfamilia.Visible = False
End Sub

Private Sub Lstfamilia_LostFocus()
Lstfamilia.Visible = False
End Sub




Private Sub Cmdafam_Click(Index As Integer)
Dim Antes
Select Case Index
Case 0
    cCve = InputBox("Introduzca la descripción de la línea a buscar", "Busqueda de Línea")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    cCve = UCase(cCve)
    Antes = DataGrid2.Bookmark 'grid de familias
    Adolinasig.Recordset.MoveFirst
    Adolinasig.Recordset.Find "Sfdescrip like '" & Trim(cCve) & "*'"
    If Adolinasig.Recordset.EOF Then
        MsgBox "La Descripcion no existe en el Catalogo "
        DataGrid2.Bookmark = Antes
    End If
Case 1
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
    adofamilia.Refresh
Case 2
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = True
    adofamilia.Refresh
    
Case 3
    Antes = DataGrid2.Bookmark 'grid de familias
    DataGrid2.Columns(2).Locked = False
    DataGrid2.Refresh
    DataGrid2.Bookmark = Antes
End Select

End Sub

