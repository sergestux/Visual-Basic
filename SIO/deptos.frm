VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form fdeptos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogo de  Departamentos en tiendas"
   ClientHeight    =   5340
   ClientLeft      =   1635
   ClientTop       =   2730
   ClientWidth     =   10830
   Icon            =   "deptos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   10455
      Begin VB.CommandButton Cmdfamilia 
         Caption         =   "&Catalogo de Familias"
         Height          =   495
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmdvlinea 
         Caption         =   "&Familias del Departamento"
         Height          =   495
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7560
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Height          =   495
         Left            =   2040
         Picture         =   "deptos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ir al ultimo registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   1440
         Picture         =   "deptos.frx":05B4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ir al siguiente registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   840
         Picture         =   "deptos.frx":0726
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ir al registro anterior"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   240
         Picture         =   "deptos.frx":0898
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ir al primer registro"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdcancela 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   8520
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin MSAdodcLib.Adodc adodeptos 
         Height          =   450
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
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
         RecordSource    =   "select * from departamento order by depclave"
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
         Height          =   495
         Left            =   9600
         Picture         =   "deptos.frx":0A0A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir del modulo"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdmodifica 
         Caption         =   "&Modificar"
         Height          =   495
         Left            =   5520
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   495
         Left            =   6480
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin MSAdodcLib.Adodc Adofamilias 
         Height          =   375
         Left            =   1560
         Top             =   0
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "familias"
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
   Begin MSAdodcLib.Adodc Adousuario 
      Height          =   375
      Left            =   240
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc Adolinea 
      Height          =   330
      Left            =   120
      Top             =   8160
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.ComboBox Cmbusuario 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   3120
         Width           =   5175
      End
      Begin VB.TextBox txtobserva 
         Height          =   1575
         Left            =   2280
         TabIndex        =   10
         Top             =   1320
         Width           =   6975
      End
      Begin VB.TextBox txtdesc 
         Height          =   375
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   8
         Top             =   720
         Width           =   6975
      End
      Begin VB.TextBox txtclave 
         Height          =   405
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Comprador"
         Height          =   375
         Left            =   1080
         TabIndex        =   27
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones:"
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Clave:"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   34
      Top             =   5010
      Width           =   10830
      _ExtentX        =   19103
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
   Begin VB.Frame Frame3 
      Caption         =   "Familias Asignadas  "
      Height          =   4695
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   9975
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "deptos.frx":0B7C
         Height          =   3975
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1.5
         RowHeight       =   15
         TabAction       =   1
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "fclave"
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
            DataField       =   "fdescrip"
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
            DataField       =   "FDEPTO"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
               DividerStyle    =   0
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1514.835
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   5400
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Cmdlinea 
         Caption         =   "&Reporte"
         Height          =   615
         Index           =   1
         Left            =   8400
         Picture         =   "deptos.frx":0B93
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Cmdlinea 
         Caption         =   "&Regresar"
         Height          =   615
         Index           =   0
         Left            =   8400
         Picture         =   "deptos.frx":10C5
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Asignar Familias   "
      Height          =   4695
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   10455
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "deptos.frx":1237
         Height          =   3255
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "fclave"
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
            DataField       =   "fdescrip"
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
            DataField       =   "fdepto"
            Caption         =   "Departamento"
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
         Caption         =   "&Cambio Depto."
         Height          =   550
         Index           =   2
         Left            =   5880
         Picture         =   "deptos.frx":1251
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3840
         Width           =   1215
      End
      Begin VB.ListBox Lstfamilia 
         Height          =   450
         ItemData        =   "deptos.frx":18BB
         Left            =   7680
         List            =   "deptos.frx":18BD
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Cmdafam 
         Caption         =   "&Regresar"
         Height          =   550
         Index           =   1
         Left            =   8640
         Picture         =   "deptos.frx":18BF
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Cmdafam 
         Caption         =   "&Buscar"
         Height          =   550
         Index           =   0
         Left            =   7320
         Picture         =   "deptos.frx":1A31
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "PARA ASIGNAR UNA FAMILIA A ESTE DEPARTAMENTO MODIFICAR LA CLAVE DE LA COLUMNA  "" Departamento """
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
         Left            =   360
         TabIndex        =   32
         Top             =   3960
         Width           =   5175
      End
      Begin VB.Label Label6 
         Caption         =   "PERTENECE AL"
         Height          =   255
         Left            =   7320
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "CATALOGO DE FAMILIAS"
         Height          =   375
         Left            =   840
         TabIndex        =   30
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "fdeptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lnuevo As Boolean


Private Sub Cmdafam_Click(Index As Integer)
Dim Antes
Select Case Index
Case 0
    cCve = InputBox("Introduzca la DESCRIPCION de la FAMILIA a buscar", "Busqueda de Familia")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    cCve = UCase(cCve)
    Antes = DataGrid2.Bookmark 'grid de familias
    Adofamilias.Recordset.MoveFirst
    Adofamilias.Recordset.Find "fdescrip like '" & Trim(cCve) & "*'"
    If Adofamilias.Recordset.EOF Then
        MsgBox "La Descripcion no existe en el Catalogo "
        DataGrid2.Bookmark = Antes
    End If
Case 1
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
    Adolinea.Refresh
Case 2
    Antes = DataGrid2.Bookmark 'grid de familias
    DataGrid2.Columns(2).Locked = False
    DataGrid2.Refresh
    DataGrid2.Bookmark = Antes
End Select

End Sub

Private Sub cmdcancela_Click()
On Error GoTo Error:
    adodeptos.Refresh
    Call asigna
    Call habilitar
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Cmdfamilia_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True

End Sub

Private Sub Cmdlinea_Click(Index As Integer)
Select Case Index
Case 0
Frame1.Visible = True
Frame2.Visible = True
Frame4.Visible = False
Frame3.Visible = False
Case 1
fmenu.CR1.WindowTitle = "Departamentos "
fmenu.CR1.ReportFileName = App.Path & "\depfamlin.rpt"
fmenu.CR1.Formulas(1) = "FORMSELEC = {DEPARTAMENTO.depclave} ='" & Trim(txtclave.Text) & "'"
fmenu.CR1.Connect = strconnect
fmenu.CR1.WindowState = crptMaximized
fmenu.CR1.Action = 1

End Select
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
 
 lnuevo = True
 
  Call nuevodep
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

Private Sub Cmdvlinea_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame3.Visible = True
End Sub

Private Sub Command1_Click()
On Error GoTo Error:

If adodeptos.Recordset.EOF = False And adodeptos.Recordset.BOF = False Then

adodeptos.Recordset.MoveFirst
Call asigna

End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command2_Click()
On Error GoTo Error:
Dim reg As Integer
reg = adodeptos.Recordset.AbsolutePosition
If reg > 1 Then
adodeptos.Recordset.MovePrevious
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

reg = adodeptos.Recordset.AbsolutePosition
treg = adodeptos.Recordset.RecordCount
If reg < treg Then

adodeptos.Recordset.MoveNext
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command4_Click()
On Error GoTo Error:
If adodeptos.Recordset.EOF = False And adodeptos.Recordset.BOF = False Then
adodeptos.Recordset.MoveLast
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command5_Click()
On Error GoTo Error:
Dim nreg As Integer
Dim reg As Integer
Dim SCOMPRADOR As String
SCOMPRADOR = " "
If Trim(Cmbusuario.Text) <> "" Then
N = InStr(1, Cmbusuario.Text, "[")
SCOMPRADOR = Mid(Cmbusuario.Text, N + 1, Len(Cmbusuario.Text) - N - 1)
End If

If Trim(txtdesc.Text) <> "" Then
If lnuevo Then adodeptos.Recordset.AddNew
adodeptos.Recordset!depclave = Trim(txtclave.Text)
adodeptos.Recordset!depdescrip = Trim(txtdesc.Text)
adodeptos.Recordset!depobserva = Trim(txtobserva.Text)
adodeptos.Recordset!depcomprador = SCOMPRADOR
adodeptos.Recordset.Update

Call habilitar
Else
    MsgBox "Favor de completar los datos ..."
    txtdesc.SetFocus
    
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



Private Sub Form_Load()
On Error GoTo Error:

Adofamilias.CommandType = adCmdText
Adofamilias.CursorType = adOpenKeyset
Adofamilias.LockType = adLockOptimistic
Adofamilias.ConnectionString = strconnect
Adofamilias.RecordSource = "select * from familias order by fdepto,fdescrip"
Adofamilias.Refresh


adodeptos.ConnectionString = strconnect
adodeptos.Refresh

Cmbusuario.Clear

  Adousuario.CommandType = adCmdText
  Adousuario.CursorType = adOpenKeyset
  Adousuario.RecordSource = "select * from usuarios order by name"
  Adousuario.ConnectionString = strconnect
  Adousuario.Refresh
    
  Do While Not Adousuario.Recordset.EOF = True
               If Not IsNull(Adousuario.Recordset!Name) Then
                Cmbusuario.AddItem Adousuario.Recordset!Name + "  [" + Trim(Str(Adousuario.Recordset!clave)) + "]"
               End If
              Adousuario.Recordset.MoveNext
            Loop
           Adousuario.Recordset.MoveFirst

Call asigna


Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub asigna()
Dim nclave As String
On Error GoTo Error:
If adodeptos.Recordset.EOF = False And adodeptos.Recordset.BOF = False Then

txtclave.Text = IIf(Not IsNull(adodeptos.Recordset!depclave), adodeptos.Recordset!depclave, "")
txtdesc.Text = IIf(Not IsNull(adodeptos.Recordset!depdescrip), adodeptos.Recordset!depdescrip, "")
txtobserva.Text = IIf(Not IsNull(adodeptos.Recordset!depobserva), adodeptos.Recordset!depobserva, "")
nclave = IIf(Not IsNull(adodeptos.Recordset!depcomprador), adodeptos.Recordset!depcomprador, "X")
lnuevo = False

Adolinea.CommandType = adCmdText
Adolinea.CursorType = adOpenKeyset
Adolinea.LockType = adLockOptimistic
Adolinea.ConnectionString = cn.ConnectionString
Adolinea.RecordSource = "select * from familias where fDEPTO = '" & Trim(txtclave.Text) & "' ORDER BY FDESCRIP"
Adolinea.Refresh
Lstfamilia.Clear
Lstfamilia.AddItem "                 "
Lstfamilia.AddItem Trim(txtdesc.Text)
Adousuario.Recordset.MoveFirst
Adousuario.Recordset.Find " clave = '" & Trim(nclave) & "'"
If Adousuario.Recordset.EOF = False Then
  Cmbusuario.Text = Adousuario.Recordset!Name + "  [" + Trim(Str(Adousuario.Recordset!clave)) + "]"
Else
   Cmbusuario.Text = ""
End If
End If

'If Trim(txtdesc.Text) <> "" Then
'    Lbldeptox = Trim(txtdesc.Text)
'Else
'    Lbldeptox = ""
'End If
'Lbldeptox.Refresh
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub nuevodep()
On Error GoTo Error:
If adodeptos.Recordset.EOF = False And adodeptos.Recordset.BOF = False Then
    adodeptos.Recordset.MoveLast
    txtclave.Text = Right("0000" + Trim(Str(Val(adodeptos.Recordset!depclave) + 1)), 3)
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
 Command1.Enabled = False
 Command2.Enabled = False
 Command3.Enabled = False
 Command4.Enabled = False
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
 Command1.Enabled = True
 Command2.Enabled = True
 Command3.Enabled = True
 Command4.Enabled = True
Exit Sub
Error:
MsgBox Err.Description
End Sub



Private Sub Lstfamilia_DblClick()
Dim Antes
Antes = DataGrid2.Bookmark
If Trim(Lstfamilia.Text) = "" Then
    Adofamilias.Refresh
    DataGrid2.Bookmark = Antes
    
Else
    DataGrid2.Columns(2).Text = Trim(txtclave.Text)
End If
Lstfamilia.Visible = False
End Sub

Private Sub Lstfamilia_LostFocus()
Lstfamilia.Visible = False
End Sub



