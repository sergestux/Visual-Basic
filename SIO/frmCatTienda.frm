VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmCatTienda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de tiendas"
   ClientHeight    =   3855
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmCatTienda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      DataField       =   "folcredito"
      Height          =   285
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   27
      ToolTipText     =   $"frmCatTienda.frx":0442
      Top             =   1920
      Width           =   855
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   26
      Top             =   3510
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "Esc=Cierra forma    Repag=Ant   Avpag=Siguiente    Inicio=Primero    Fin=Ultimo"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   5850
      TabIndex        =   18
      Top             =   2430
      Width           =   5850
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4560
         TabIndex        =   23
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   300
         Left            =   3480
         TabIndex        =   22
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edición"
         Height          =   300
         Left            =   1320
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ag&regar"
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         Height          =   300
         Left            =   1560
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   3120
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   5850
      TabIndex        =   12
      Top             =   3090
      Width           =   5850
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   5280
         Picture         =   "frmCatTienda.frx":04E8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4920
         Picture         =   "frmCatTienda.frx":05EA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   480
         Picture         =   "frmCatTienda.frx":06EC
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   120
         Picture         =   "frmCatTienda.frx":07EE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "folcontado"
      Height          =   285
      Index           =   5
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   $"frmCatTienda.frx":08F0
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "direccion"
      Height          =   285
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   70
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tizona"
      Height          =   285
      Index           =   3
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1500
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "telefonos"
      Height          =   285
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1185
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tidescrip"
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   495
      Width           =   4215
   End
   Begin VB.TextBox txtClave 
      DataField       =   "ticlave"
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   1
      Top             =   180
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Folio crédito"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   28
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Folio contado"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Direccion:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   855
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Zona:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Teléfonos:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1185
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripción:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   495
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Clave:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "frmCatTienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub CmdBuscar_Click()
On Error GoTo eRROR:
Dim marca
marca = adoPrimaryRS.Bookmark
ctie = InputBox("Introduzca la descripcion de la tienda", "Solicitando tienda")
adoPrimaryRS.MoveFirst
adoPrimaryRS.Find "TIDESCRIP LIKE '" & ctie & "*'"
If adoPrimaryRS.EOF Then
   MsgBox "LA DESCRIPCION DE LA TIENDA NO SE ENCUENTRA EN EL CATALOGO", vbInformation
   adoPrimaryRS.Bookmark = marca
End If
Exit Sub
eRROR:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open cCadConex
  Set adoPrimaryRS = New Recordset
  'adoPrimaryRS.Open "select ticlave,tidescrip,tiresponsable,tizona,direccion,prioridad,server from CATTIENDA Order by prioridad", db, adOpenStatic, adLockOptimistic
  adoPrimaryRS.Open "SELECT * FROM CATTIENDA Order by prioridad", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtfields
    Set oText.DataSource = adoPrimaryRS
  Next
  Set txtclave.DataSource = adoPrimaryRS
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.eRROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  lblStatus.Caption = "Registro: " & CStr(adoPrimaryRS.AbsolutePosition) & " de " & CStr(adoPrimaryRS.RecordCount)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
Dim rsttemp As New ADODB.Recordset
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT MAX(CONVERT(INT,TICLAVE)) AS UltClave FROM CATTIENDA", cn, adOpenDynamic, adLockReadOnly, adCmdText
    txtclave.Text = rsttemp!UltClave + 1

    lblStatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub


Private Sub cmdEdit_Click()
'  On Error GoTo EditErr

  lblStatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  CmdBuscar.Visible = bVal
  Cmdnext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  Dim oText As TextBox
  'Bloquea / Desbloquea los cuadros de texto
  For Each oText In Me.txtfields
     oText.Locked = bVal
  Next

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)

End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub
