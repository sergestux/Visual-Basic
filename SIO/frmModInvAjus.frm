VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmModInvAjus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del  ajuste a inventario"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "frmModInvAjus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMerma 
      Caption         =   "&Merma"
      DataField       =   "a_merma"
      DataSource      =   "AdoAjustes"
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
      Left            =   4320
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cmbUsuario 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc AdoAjustes 
      Height          =   330
      Left            =   -120
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
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
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoAjustes"
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
   Begin MSACAL.Calendar Cal1 
      Height          =   1935
      Left            =   1560
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
      _Version        =   524288
      _ExtentX        =   4683
      _ExtentY        =   3413
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2000
      Month           =   6
      Day             =   14
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   2
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3600
      Picture         =   "frmModInvAjus.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   1320
      Picture         =   "frmModInvAjus.frx":05B4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "a_observaciones"
      DataSource      =   "AdoAjustes"
      Height          =   615
      Index           =   3
      Left            =   1680
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "a_motivo"
      DataSource      =   "AdoAjustes"
      Height          =   525
      Index           =   2
      Left            =   1680
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "a_fecha"
      DataSource      =   "AdoAjustes"
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "a_responsable"
      DataSource      =   "AdoAjustes"
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Motivo"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Fecha de ajuste"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Responsable"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmModInvAjus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsttemp As ADODB.Recordset
Private Sub Cal1_Click()
txtcampos(1).Text = Cal1.Value
Cal1.Visible = False
txtcampos(1).SetFocus
SendKeys "{TAB}"
End Sub

Private Sub Cal1_LostFocus()
txtcampos(1).Text = Cal1.Value
Cal1.Visible = False
End Sub

Private Sub cmbUsuario_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub cmbusuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii Then
   KeyAscii = 0
   SendKeys "{Tab}"
End If

End Sub

'Private Sub cmbusuario_Validate(Cancel As Boolean)
'Dim n As Integer
'rsttemp.MoveFirst
'rsttemp.Find "NAME = '" & cmbUsuario.Text & "'"
'If rsttemp.EOF = True Then
'   MsgBox "Debe seleccionar un usuario de la lista desplegable", vbExclamation
'   cmbUsuario.SetFocus
'   Cancel = True
'Else
'   txtCampos(0).Text = rsttemp!Clave
'   txtCampos(0).SetFocus
'End If
'End Sub

Private Sub cmdCancelar_Click()
  frmModInv.dbgrdModInv.AllowUpdate = False
  frmModInv.CmdActualizar.Enabled = True
  nOp = 5
  Unload Me
End Sub

Private Sub cmdGrabar_Click()
Dim rstDetAju As ADODB.Recordset
On Error GoTo Error
  lDatAJu = Trim(txtcampos(0).Text)
  AdoAjustes.Recordset.Update
  nOp = AdoAjustes.Recordset!a_clave
  Set rstDetAju = New ADODB.Recordset
  rstDetAju.Open "DETALLEAJUSTES", cCadConex, adOpenKeyset, adLockOptimistic, adCmdTable
  
  Set frmModInv.AdoDetAju.Recordset = rstDetAju
  frmModInv.AdoDetAju.Refresh
  Unload Me
'  SendKeys "{DOWN}"
  Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     KeyAscii = 0
     SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
'Set rsttemp = New ADODB.Recordset
'rsttemp.Open "USUARIOS", cCadConex, adOpenDynamic, adLockOptimistic, adCmdTable
'While Not rsttemp.EOF
'   cmbUsuario.AddItem rsttemp!Name
'   rsttemp.MoveNext
'Wend

AdoAjustes.ConnectionString = cCadConex
AdoAjustes.CommandType = adCmdTable
AdoAjustes.RecordSource = "AJUSTES"
AdoAjustes.Refresh
AdoAjustes.Recordset.AddNew

'txtCampos(0).Text = Trim(Mid(cCveDesUsu, 1, 3))
'cmbUsuario.Text = Trim(Mid(cCveDesUsu, 3))
txtcampos(0).Enabled = False
Cmbusuario.Enabled = False
txtcampos(1).Text = date + Time
txtcampos(1).Enabled = False
End Sub

Private Sub txtCampos_GotFocus(Index As Integer)
Select Case Index
   Case 1  'Campos correspondientes a las fechas
        If txtcampos(Index).Text = "" Then
          Cal1.Value = Trim(Str(Month(date))) + "/" + Trim(Str(Day(date))) + "/" + Mid(Trim(Str(Year(date))), 3, 2)
          Cal1.Refresh
        Else
          Cal1.Value = txtcampos(Index).Text
        End If
        Cal1.Visible = True
        'Cal1.Top = txtCampos(Index).Top + txtCampos(Index).Height + 600
        'Cal1.Left = txtCampos(Index).Left - 1000
   Case Else
        Cal1.Visible = False
End Select
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
Select Case Index
    Case 2
        If IsNull(txtcampos(Index).Text) Or Trim(txtcampos(Index).Text) = "" Then
           MsgBox "Es necesario especificar un motivo ", vbExclamation
           txtcampos(Index).SetFocus
        End If
End Select
End Sub
