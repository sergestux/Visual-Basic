VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmCatusu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de usuarios"
   ClientHeight    =   4230
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmCatUsu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      DataField       =   "caja"
      Height          =   285
      Index           =   7
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   33
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "permisos"
      Height          =   285
      Index           =   6
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   31
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox CHKCAJERO 
      Caption         =   "Cajero"
      Height          =   375
      Left            =   4800
      TabIndex        =   30
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "level1"
      Height          =   285
      Index           =   5
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   28
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "passmaster"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   25
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6330
      TabIndex        =   16
      Top             =   2940
      Width           =   6330
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   5280
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   300
         Left            =   4320
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   2640
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edición"
         Height          =   300
         Left            =   960
         TabIndex        =   18
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ag&regar"
         Height          =   300
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   3240
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         Height          =   300
         Left            =   1560
         TabIndex        =   22
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
      ScaleWidth      =   6330
      TabIndex        =   10
      Top             =   3465
      Width           =   6330
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   5280
         Picture         =   "frmCatUsu.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4920
         Picture         =   "frmCatUsu.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   480
         Picture         =   "frmCatUsu.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   120
         Picture         =   "frmCatUsu.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   11
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
         TabIndex        =   15
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtClave 
      DataField       =   "Clave"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "pass"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtFields 
      DataField       =   "sucursal"
      Height          =   285
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1905
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "name"
      Height          =   285
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "login"
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   1
      Top             =   540
      Width           =   1335
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   29
      Top             =   3885
      Width           =   6330
      _ExtentX        =   11165
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
   Begin VB.Label lblLabels 
      Caption         =   "Comprador"
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   34
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Permisos"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   32
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Departamento"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Clave maestra"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Clave:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   225
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Contraseña:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1215
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sucursal:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1905
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nombre:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   855
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Login:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "frmCatusu"
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
ctie = InputBox("Introduzca el nombre del usuario", "Solicitando usuario")
adoPrimaryRS.MoveFirst
adoPrimaryRS.Find "NAME LIKE '*" & ctie & "*'"
If adoPrimaryRS.EOF Then
   MsgBox "LA DESCRIPCION DEL USUARIO NO SE ENCUENTRA EN EL CATALOGO", vbInformation
   adoPrimaryRS.Bookmark = marca
End If
If adoPrimaryRS!LEVEL1 = "J" Then
     CHKCAJERO.Value = 1
  Else
     CHKCAJERO.Value = 0
  End If
Exit Sub
eRROR:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
Dim rsttemp As ADODB.Recordset
Set rsttemp = New ADODB.Recordset
rsttemp.Open "SELECT * FROM USUARIOS", cn, adOpenKeyset, adLockOptimistic
While Not rsttemp.EOF
   clave = ""
   For N = 1 To Len(Trim(rsttemp!pass))
      clave = clave + Chr(Asc(Mid(Trim(rsttemp!pass), N, 1)) + 30)
   Next
   rsttemp!pass = clave
   rsttemp.Update
   rsttemp.MoveNext
Wend
End Sub

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open cCadConex
  Set adoPrimaryRS = New Recordset
  If lpprod Then
        adoPrimaryRS.Open "Select * from Usuarios where level1 = 'G' or level1 = 'J' Order by NAME  ", db, adOpenStatic, adLockOptimistic
  Else
        adoPrimaryRS.Open "Select * from Usuarios Order by NAME", db, adOpenStatic, adLockOptimistic
  End If
  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtfields
    Set oText.DataSource = adoPrimaryRS
  Next
  Set txtclave.DataSource = adoPrimaryRS
'  Set txtSucur.DataSource = adoPrimaryRS
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
On Error Resume Next
  'Esto mostrará la posición de registro actual para este Recordset
  lblStatus.Caption = "Registro: " & CStr(adoPrimaryRS.AbsolutePosition) & " de " & CStr(adoPrimaryRS.RecordCount)
  'ClaveD = ""
  'For n = 1 To Len(Trim(adoPrimaryRS!Pass))
  '    ClaveD = ClaveD + Chr(Asc(Mid(Trim(adoPrimaryRS!Pass), n, 1)) - 30)
  'Next
  'txtFields(3).Text = ClaveD
' MsgBox ClaveD
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
Dim rsttemp As ADODB.Recordset
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT MAX(CONVERT(INT,CLAVE)) AS UltClave FROM USUARIOS", cn, adOpenDynamic, adLockReadOnly, adCmdText
    txtclave.Text = rsttemp!UltClave + 1
    txtfields(0).Text = Trim(Mid(cSucursal, 1, 3))
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
  If lpprod Then
     adoPrimaryRS!LEVEL1 = "G"
     adoPrimaryRS!pass = "XX"
  End If
  If CHKCAJERO.Value = 1 Then
     adoPrimaryRS!LEVEL1 = "J"
     adoPrimaryRS!pass = "XX"
  End If
  If mbAddNewFlag Then
    'validacion del nombre del usuario
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'MsgBox txtfields(1).Text
    rs.Open "SELECT * FROM USUARIOS where login = '" & Trim(txtfields(1).Text) & "'", cn, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
       MsgBox "Este usuario ya fue dado de alta, utilice otro LOGIN ", vbInformation, "SISTEMA INTEGRAL PITICO"
       adoPrimaryRS.Delete
       'adoPrimaryRS.CancelUpdate
       Exit Sub
    End If
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If
  adoPrimaryRS.UpdateBatch adAffectAll

  
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
  If adoPrimaryRS!LEVEL1 = "J" Then
     CHKCAJERO.Value = 1
  Else
     CHKCAJERO.Value = 0
  End If
  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False
  If adoPrimaryRS!LEVEL1 = "J" Then
     CHKCAJERO.Value = 1
  Else
     CHKCAJERO.Value = 0
  End If
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
  If adoPrimaryRS!LEVEL1 = "J" Then
     CHKCAJERO.Value = 1
  Else
     CHKCAJERO.Value = 0
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
  If adoPrimaryRS!LEVEL1 = "J" Then
     CHKCAJERO.Value = 1
  Else
     CHKCAJERO.Value = 0
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
  'Bloquea / Desbloquea los cuadros de texto 'Eric
  For Each oText In Me.txtfields
     oText.Locked = bVal
  Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
lpprod = False
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
   Exit Sub
End If
If Index = 3 Then 'Campo correspondiente a contraseña
   CarValido = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
   If InStr(1, CarValido, Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
      MsgBox "Para especificar la contraseña solo se permiten numeros y letras", vbInformation
      KeyAscii = 0
   End If
End If
End Sub

Private Sub txtfields_LostFocus(Index As Integer)
'If Index = 3 Then
'   MsgBox txtFields(Index).Text
'   Clave = ""
'   For n = 1 To Len(Trim(txtFields(Index).Text))
'      Clave = Clave + Chr(Asc(Mid(Trim(txtFields(Index).Text), n, 1)) + 30)
'   Next
'   MsgBox Clave
'   ClaveD = ""
'   For n = 1 To Len(Trim(Clave))
'      ClaveD = ClaveD + Chr(Asc(Mid(Trim(Clave), n, 1)) - 30)
'   Next
'   MsgBox ClaveD
'End If
End Sub
