VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form flineas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de Lineas"
   ClientHeight    =   4695
   ClientLeft      =   1350
   ClientTop       =   1515
   ClientWidth     =   8940
   Icon            =   "frmlineas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   4320
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "                                                                               Para salir presione la Tecla   [Esc]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   8655
      Begin MSAdodcLib.Adodc adofamilia 
         Height          =   375
         Left            =   3480
         Top             =   0
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
         RecordSource    =   "select * from lineas order by sfdescrip"
         Caption         =   "familia"
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
      Begin VB.CommandButton Cmdbus 
         Height          =   400
         Left            =   2760
         Picture         =   "frmlineas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   400
         Left            =   4680
         TabIndex        =   19
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdmodifica 
         Caption         =   "&Modificar"
         Height          =   400
         Left            =   3480
         TabIndex        =   18
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdsalir 
         Height          =   400
         Left            =   7920
         Picture         =   "frmlineas.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmdcancela 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   400
         Left            =   6840
         TabIndex        =   16
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Height          =   400
         Left            =   360
         Picture         =   "frmlineas.frx":06AE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ir al primer registro"
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Height          =   400
         Left            =   960
         Picture         =   "frmlineas.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Height          =   400
         Left            =   1560
         Picture         =   "frmlineas.frx":0992
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Height          =   400
         Left            =   2160
         Picture         =   "frmlineas.frx":0B04
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Actualizar"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5760
         TabIndex        =   11
         ToolTipText     =   "Buscar Registro"
         Top             =   200
         Width           =   975
      End
      Begin MSAdodcLib.Adodc adolineas 
         Height          =   330
         Left            =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
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
         Appearance      =   0
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
         RecordSource    =   "select fclave,fdescrip,depdescrip from familias,departamento  where fdepto=depclave "
         Caption         =   "AdoLineas"
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
   Begin VB.Frame Frame1 
      DragIcon        =   "frmlineas.frx":0C76
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtsfobserva 
         Height          =   1575
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtsfdescrip 
         Height          =   375
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   6855
      End
      Begin VB.TextBox txtsfclave 
         Height          =   405
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   3000
         Width           =   5295
      End
      Begin VB.TextBox txtsffamilia 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Departamento"
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
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Lbldepto 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Clave:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3000
         Width           =   1215
      End
   End
   Begin VB.Frame frmbusca 
      Height          =   3855
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton Cmdreg 
         Caption         =   "&Regresar"
         Height          =   255
         Left            =   7440
         TabIndex        =   24
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ListBox Lstlin 
         Height          =   2985
         Left            =   480
         TabIndex        =   21
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "flineas"
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
Adofamilia.Refresh
Do While Not Adofamilia.Recordset.EOF = True
   If Not IsNull(Adofamilia.Recordset!sfdescrip) Then
      Lstlin.AddItem Adofamilia.Recordset!sfdescrip + "  [" + Adofamilia.Recordset!sfclave + "]"
   End If
   Adofamilia.Recordset.MoveNext
Loop
Adofamilia.Recordset.MoveFirst
stb1.SimpleText = Space(35) & "Para seleccionar la linea dar doble click sobre ella"

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdcancela_Click()
On Error GoTo Error:
    Adofamilia.Refresh
    Call asigna
    Call habilitar
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdmodifica_Click()
On Error GoTo Error:
lnuevo = False
txtsfclave.Locked = True
txtsfdescrip.SetFocus
Call dhabilitar

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdnuevo_Click()
On Error GoTo Error:
 txtsfdescrip.Text = ""
 txtsfobserva.Text = ""
 txtsffamilia.Text = ""
 Lbldepto.Caption = ""
 
 Combo2.Text = ""
 lnuevo = True
  Call nuevalin
  Call dhabilitar
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Cmdreg_Click()
Frame1.Visible = True
Frame2.Visible = True
frmbusca.Visible = False

End Sub

Private Sub cmdsalir_Click()
On Error GoTo Error:
Unload Me
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Combo2_Click()
On Error GoTo Error:
         If Trim(Combo2.Text) <> "" Then
         SendKeys "{TAB}"
         End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command6_Click()
On Error GoTo Error:
If Adofamilia.Recordset.EOF = False And Adofamilia.Recordset.BOF = False Then
Adofamilia.Recordset.MoveFirst
Call asigna
End If

Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command2_Click()
On Error GoTo Error:
Dim reg As Integer
reg = Adofamilia.Recordset.AbsolutePosition
If reg > 1 Then
Adofamilia.Recordset.MovePrevious
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

reg = Adofamilia.Recordset.AbsolutePosition
treg = Adofamilia.Recordset.RecordCount
If reg < treg Then

Adofamilia.Recordset.MoveNext
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command4_Click()
On Error GoTo Error:
If Adofamilia.Recordset.EOF = False And Adofamilia.Recordset.BOF = False Then
Adofamilia.Recordset.MoveLast
Call asigna
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command5_Click()
On Error GoTo Error:
If Trim(txtsfdescrip.Text) <> "" And Trim(txtsffamilia.Text) <> "" Then
If lnuevo Then
    Adofamilia.Recordset.AddNew
End If
Adofamilia.Recordset!sfclave = txtsfclave.Text
Adofamilia.Recordset!sfdescrip = txtsfdescrip.Text
Adofamilia.Recordset!sfobserva = txtsfobserva.Text
Adofamilia.Recordset!sffamilia = txtsffamilia.Text


Adofamilia.Recordset.Update
Call habilitar
Else
    MsgBox "Favor de completar los datos ..."
    txtsfdescrip.SetFocus
    
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



Private Sub asigna()
'On Error GoTo error:
If Adofamilia.Recordset.EOF = False And Adofamilia.Recordset.BOF = False Then

txtsfclave.Text = Adofamilia.Recordset!sfclave
txtsfdescrip.Text = Adofamilia.Recordset!sfdescrip
txtsfobserva.Text = Adofamilia.Recordset!sfobserva
txtsffamilia.Text = Adofamilia.Recordset!sffamilia
txtsffamilia_KeyPress 13
lnuevo = False

End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub nuevalin()
Dim rs As ADODB.Recordset
On Error GoTo Error:
If Adofamilia.Recordset.EOF = False And Adofamilia.Recordset.BOF = False Then
    Set rs = New ADODB.Recordset
    rs.Open "SELECT max(sfclave) AS NvaCve from lineas", cn, adOpenDynamic, adLockOptimistic, adCmdText
    'Adofamilia.Recordset.MoveLast
    'txtsfclave.Text = Right("0000" + Trim(Str(Val(Adofamilia.Recordset!sfclave) + 1)), 3)
    txtsfclave.Text = Right("0000" + Trim(Str(rs!NvaCve + 1)), 3)
    txtsfclave.Locked = True
    
    txtsfdescrip.SetFocus
        
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
 Exit Sub
Error:
MsgBox Err.Description
 End Sub



Private Sub cmdgraba_Click()
On Error GoTo Error:
Adofamilia.Recordset!fdepto = txtfdepto.Text
Adofamilia.Recordset.Update
Adofamilia.Refresh
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
Adofamilia.Recordset.MoveFirst
Adofamilia.Recordset.Find "sfclave = '" & Cvelin & "'"
If Adofamilia.Recordset.EOF Then
   MsgBox "   "
End If
frmbusca.Visible = False
Frame1.Visible = True
Frame2.Visible = True
stb1.SimpleText = Space(45) + "Para salir presione la Tecla   [Esc]"
stb1.Refresh

Call asigna

End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtsfclave_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
            If Trim(txtclave.Text) <> "" Then
          
             Adofamilia.Recordset.MoveFirst
             Adofamilia.Recordset.Find "fclave = '" & Trim(txtclave.Text) & "'"
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
             adolineas.Recordset.MoveFirst
             adolineas.Recordset.Find "fclave = '" & Cvedepto & "'"
             If adolineas.Recordset.EOF Then
                MsgBox "Seleccione un departamento "
                Combo2.SetFocus
             Else
             txtsffamilia.Text = adolineas.Recordset!fclave
             Lbldepto.Caption = Trim(adolineas.Recordset!depdescrip)
             
             
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
'On Error GoTo error:

Adofamilia.ConnectionString = strconnect
adolineas.ConnectionString = strconnect
Adofamilia.Refresh

Combo2.Clear
    adolineas.Refresh
    Do While Not adolineas.Recordset.EOF = True
               If Not IsNull(adolineas.Recordset!depdescrip) Then
                 Combo2.AddItem adolineas.Recordset!fdescrip + "  [" + adolineas.Recordset!fclave + "]"
               End If
              adolineas.Recordset.MoveNext
            Loop
           adolineas.Recordset.MoveFirst

Call asigna
Exit Sub
Error:
MsgBox Err.Description
           
End Sub


Private Sub txtsffamilia_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
            If Trim(txtsffamilia.Text) <> "" Then
             adolineas.Refresh
             adolineas.Recordset.MoveFirst
             adolineas.Recordset.Find "fclave = '" & Trim(txtsffamilia.Text) & "'"
             If adolineas.Recordset.EOF Then
                MsgBox "Seleccione un departamento "
                Combo2.SetFocus
             Else
                Combo2.Text = adolineas.Recordset!fdescrip + "[" + adolineas.Recordset!fclave + "]"
                txtsffamilia.Text = adolineas.Recordset!fclave
                Lbldepto.Caption = Trim(adolineas.Recordset!depdescrip)
                      
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

