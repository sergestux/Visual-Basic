VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmInvent 
   Caption         =   "Captura de inventario"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   Icon            =   "frmInvent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport cr1 
      Left            =   4680
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowLeft      =   0
      WindowTop       =   0
      WindowState     =   2
   End
   Begin VB.CommandButton cmdopcion 
      Caption         =   "Fin."
      Height          =   615
      Index           =   2
      Left            =   4200
      Picture         =   "frmInvent.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Guarda existencias actuales"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdopcion 
      Caption         =   "Resp"
      Height          =   615
      Index           =   1
      Left            =   3720
      Picture         =   "frmInvent.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Guarda existencias actuales"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdopcion 
      Caption         =   "Zona"
      Height          =   615
      Index           =   0
      Left            =   3240
      Picture         =   "frmInvent.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Existencias de productros por Zona"
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox cmbZona 
      Height          =   315
      ItemData        =   "frmInvent.frx":0B78
      Left            =   1680
      List            =   "frmInvent.frx":0BA9
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2880
      Picture         =   "frmInvent.frx":0BE0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      Picture         =   "frmInvent.frx":0D52
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtPieza 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtCaja 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtclave 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Zona"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Código"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Piezas"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Cajas"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblProducto 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5055
   End
End
Attribute VB_Name = "frmInvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbZona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   txtclave.SetFocus
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub cmbZona_LostFocus()
If Not IsNumeric(Me.cmbZona.Text) Then
   MsgBox "Debe seleccionar una zona de la lista desplegable", vbInformation
   cmbZona.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub cmdGrabar_Click()
If Not IsNumeric(cmbZona.Text) Then
   MsgBox "Es necesario especificar la zona de captura", vbInformation, "Zona"
   cmbZona.SetFocus
   Exit Sub
End If
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM InvInicial WHERE clave = '" & Trim(txtclave.Text) & "' And Zona = " & cmbZona.Text, cn, adOpenForwardOnly, adLockOptimistic, admcdtext
If rs.BOF And rs.EOF Then
   rs.AddNew
   rs!clave = txtclave.Text
   rs!cajas = txtCaja.Text
   rs!piezas = txtPieza.Text
   rs!ZONA = cmbZona.Text
   rs!USUARIO = Trim(cUsuario)
Else
   rs!cajas = txtCaja.Text
   rs!piezas = txtPieza.Text
   rs!ZONA = cmbZona.Text
End If
rs.Update
rs.Close
Set rs = Nothing
txtclave.Text = ""
txtCaja.Text = ""
txtPieza.Text = ""
txtclave.SetFocus
cmdGrabar.Enabled = False
End Sub

Private Sub cmdopcion_Click(Index As Integer)
Select Case Index
    Case 0
        CR1.ReportFileName = App.Path & "\InvIni.RPT"
        CR1.WindowTitle = "Captura de inventario de la zona " & cmbZona.Text
        CR1.Connect = cCadConex
         CR1.Formulas(0) = "FORMSELEC = " & Me.cmbZona.Text
         CR1.Action = 1
    Case 1
         If MsgBox("Deseas respaldar existencias actuales y poner en ceros el inventario", vbQuestion + vbYesNo, "Incializar inventario") = vbYes Then
            cn.Execute "UPDATE inventario SET cajasant = incant, pzasant = incantpza"
            'cn.Execute "UPDATE inventario SET incant = 0, incantpza = 0, ininicial = 0, ininicialp = 0"
            MsgBox "La incialización de inventario se realizo correctamente", vbInformation
         End If
     Case 2
     
End Select
End Sub

Private Sub Form_Activate()
cmbZona.SetFocus
End Sub

Private Sub txtCaja_GotFocus()
  txtCaja.SelStart = 0
  txtCaja.SelLength = Len(txtCaja.Text)
End Sub

Private Sub txtCaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   txtPieza.SetFocus
End If
End Sub

Private Sub txtclave_GotFocus()
  txtclave.SelStart = 0
  txtclave.SelLength = Len(txtclave.Text)
End Sub

Private Sub txtclave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   txtCaja.SetFocus
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub txtclave_LostFocus()
If Trim(txtclave.Text) = "" Then
    MsgBox "Debe especificar el código del producto", vbInformation, "Teclee Código"
    'txtclave.SetFocus
    Exit Sub
End If
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT Descripc, LTRIM(STR(paquetes)) + ' x ' + LTRIM(STR(CONTENID,10,3)) + space(1) + RTRIM(medida) AS MEDIDA FROM TFPRODUC WHERE consec = " & txtclave.Text, cn, adLockOptimistic, adCmdText
If rs.BOF And rs.EOF Then
   MsgBox "El código del producto no existe", vbInformation, "Codigo inexistente"
   rs.Close
   Set rs = Nothing
   Exit Sub
Else
   lblProducto.Caption = rs!DESCRIPC & Chr(13) & rs!medida
End If
rs.Close
rs.Open "SELECT * FROM InvInicial WHERE clave = '" & Trim(txtclave.Text) & "' And Zona = " & cmbZona.Text, cn, adOpenForwardOnly, adLockOptimistic, admcdtext
If Not (rs.BOF And rs.EOF) Then
   txtCaja.Text = rs!cajas
   txtPieza.Text = rs!piezas
Else
   txtCaja.Text = 0
   txtPieza.Text = 0
End If
cmdGrabar.Enabled = True
rs.Close
Set rs = Nothing
End Sub


Private Sub txtPieza_GotFocus()
  txtPieza.SelStart = 0
  txtPieza.SelLength = Len(txtPieza.Text)
End Sub

Private Sub txtPieza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   cmdGrabar.SetFocus
End If
End Sub
