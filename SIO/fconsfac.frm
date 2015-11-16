VERSION 5.00
Begin VB.Form fconfac 
   Caption         =   "Validacion de Factura y Serie"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   Icon            =   "fconsfac.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraagrega 
      BackColor       =   &H80000013&
      Caption         =   "Datos Para nueva Factura"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmdcte 
         Caption         =   "Asigna Cliente"
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtaserie 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtafactura 
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton btnagregar 
         Caption         =   "&Verificar"
         Height          =   735
         Left            =   4920
         Picture         =   "fconsfac.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cmbcliente 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Serie "
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Factura Inicial"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
   End
End
Attribute VB_Name = "fconfac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnagregar_Click()
Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
CAD = "SELECT numfactura,serie FROM facventa WHERE numfactura = '" & Trim(txtafactura.Text) & "' and serie = '" & Trim(txtaserie.Text) & "'"
RsCon.Open CAD, cn, adOpenKeyset, adLockOptimistic, adCmdText
pr = InStr(cmbCliente.Text, "[")
tt = Mid(cmbCliente.Text, pr + 1, Len(cmbCliente.Text))
pr = InStr(tt, "]")
tt = Mid(tt, 1, pr - 1)
If RsCon.EOF Then
   lpprov = True
Else
   lpprov = False
End If
Me.Hide
End Sub

Private Sub cmdcte_Click()
If MsgBox("DESEAS ASIGNAR ESTE CLIENTE A LA FACTURA SELECCIONADA", vbYesNo + vbInformation) = vbYes Then
   pr = InStr(cmbCliente.Text, "[")
   tt = Mid(cmbCliente.Text, pr + 1, Len(cmbCliente.Text))
   pr = InStr(tt, "]")
   tt = Mid(tt, 1, pr - 1)
   cn.Execute "UPDATE facventa SET cancelado = 0, faccliente = " & tt & ",globconfin  = 0, RFC = CRFC FROM catcliente WHERE rtrim(serie) +'-'+ numfactura  = '" & frmFacturas.AdoFacturas.Recordset!Factura & "' AND CCLAVE = FACCLIENTE"
   'MsgBox "UPDATE facventa_det SET RFC = F.RFC FROM facventa F WHERE rtrim(serie) +'-'+ numfactura  = '" & frmFacturas.AdoFacturas.Recordset!Factura & "' AND f.serie = FACventa_det.serie AND f.numfactura = facventa_det.factura"
   cn.Execute "UPDATE facventa_det SET RFC_det = F.RFC FROM facventa F WHERE rtrim(f.serie) +'-'+ numfactura  = '" & frmFacturas.AdoFacturas.Recordset!Factura & "' AND f.serie = FACventa_det.serie AND f.numfactura = facventa_det.factura"
   lpprov = False
End If
End Sub

Private Sub Form_Activate()
Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT cclave,cnombre,crfc FROM catcliente WHERE ctipo = 0 AND len(crfc) > 10 ORDER BY cnombre", cn, adOpenKeyset, adLockOptimistic, adCmdText
If RsCon.EOF Then
           MsgBox "No existen Clientes en el Catalogo", vbInformation
           Exit Sub
Else
    While Not RsCon.EOF
       cmbCliente.AddItem RsCon!cNombre & " [" & RsCon!cclave & "]"
       RsCon.MoveNext
    Wend
End If
cmbCliente.ListIndex = 0
Set RsCon = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   lpprov = True
   Me.Hide
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Form_Load()
lpprov = False
End Sub
