VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmnewprod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificacion/captura de Productos de la Factura"
   ClientHeight    =   1980
   ClientLeft      =   450
   ClientTop       =   330
   ClientWidth     =   9960
   ControlBox      =   0   'False
   Icon            =   "Fnewprod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   9960
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   1725
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "F1 = Agregar"
            TextSave        =   "F1 = Agregar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "F2 = CAMBIAR"
            TextSave        =   "F2 = CAMBIAR"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "F3 = ELIMINAR"
            TextSave        =   "F3 = ELIMINAR"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "F4 = CERRAR"
            TextSave        =   "F4 = CERRAR"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "F5 = CANTIDADES"
            TextSave        =   "F5 = CANTIDADES"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "F6 = PRODUCTOS"
            TextSave        =   "F6 = PRODUCTOS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fracambia 
      BackColor       =   &H80000016&
      Caption         =   "Modificacion de Importes"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton Command5 
         Caption         =   "&Productos"
         Height          =   450
         Left            =   5520
         TabIndex        =   1
         Top             =   1080
         Width           =   900
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   450
         Left            =   6360
         TabIndex        =   2
         Top             =   1080
         Width           =   900
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Agregar"
         Height          =   450
         Left            =   7200
         TabIndex        =   9
         Top             =   1080
         Width           =   900
      End
      Begin VB.CommandButton btncambia 
         Caption         =   "&Cambiar"
         Height          =   450
         Left            =   8040
         TabIndex        =   3
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtprepza 
         Height          =   285
         Left            =   5760
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdRegresa 
         Caption         =   "&Cerrar"
         Height          =   450
         Left            =   8880
         TabIndex        =   18
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtcajas 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtpzas 
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtuni 
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txttot 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7800
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cmbprod 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label4 
         Caption         =   "="
         Height          =   255
         Left            =   7200
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Precio Pieza"
         Height          =   255
         Left            =   5760
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ventas 
         Caption         =   "ventas"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Factura 
         Caption         =   "Factura"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Seleccion del Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Cajas"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Piezas"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Precio Caja"
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   7800
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmnewprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btncambia_Click()
prod = frmFactDet.AdoFacDet.Recordset!producto
'SERIE = Mid(Factura.Caption, 1, 2)
'If Mid(SERIE, 2, 1) = "-" Then
'   SERIE = Mid(SERIE, 1, 1)
'End If
'If Len(SERIE) = 1 Then
'    Factura = Mid(Factura.Caption, 3, Len(Factura.Caption))
'Else
'    Factura = Mid(Factura.Caption, 4, Len(Factura.Caption))
'End If
nSepara = InStr(1, Factura.Caption, "-")
SERIE = Mid(Factura.Caption, 1, nSepara - 1)
Factura = Trim(Mid(Factura.Caption, nSepara + 1, Len(Factura.Caption)))

CAD = " update facventa_Det set cantidad = " & txtcajas.Text & " , cantidadp = " & txtpzas.Text & " , precio = " & txtuni.Text & " , importe = " & txttot.Text & " where producto = '" & Trim(prod) & "' and factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
cn.Execute CAD
RESP = MsgBox("Deseas  Actualizar Iva y Ieps ? ", vbYesNo, "CORRECCION")
If RESP = vbYes Then
   ' SE VERIFICA LA TASA QUE DEBE TENER
   If IsNull(frmFactDet.AdoFacDet.Recordset!tasaieps) Or (frmFactDet.AdoFacDet.Recordset!tasaieps < 1) Then
           'vtasaieps = frmFactDet.validatasa(frmFactDet.AdoFacDet.Recordset!ieps, AdoFacDet.Recordset!iva)
           Dim rs As ADODB.Recordset
           Set rs = New ADODB.Recordset
           rs.Open "SELECT tasaieps FORM TFPRODUC WHERE consec = '" & Trim(prod) & "'", cn, adOpenForwardOnly, adLockOptimistic, admcdtext
           vtasaieps = rs!tasaieps
           CAD = "update facventa_det set tasaieps  = " & vtasaieps & " where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
           cn.Execute CAD
   End If
   'el iva
   CAD = " update facventa set iva = " & _
         " (select sum(importe - importe / (1 + (iva/100)) ) " & _
         " from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' ) " & _
         " Where numfactura = '" & Trim(Factura) & "' And serie = '" & Trim(SERIE) & "'"
   cn.Execute CAD
   'el ieps
   CAD = "update facventa set ieps = " & _
         " (select sum( IMPORTE - (importe / (1 + (iva/100)) / (1 + (IEPS/100)) " & _
         " + (importe - importe / (1 + (iva/100)))  )  ) " & _
         " from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' )" & _
         " Where numfactura = '" & Trim(Factura) & "' And serie = '" & Trim(SERIE) & "'"
    cn.Execute CAD
End If
'cad = "update facventa set total = (select sum(importe) from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(serie) & "')  where numfactura = '" & Trim(Factura) & "' and serie = '" & Trim(serie) & "'"
'cn.Execute cad
'SE ACTUALIZA EL MONTO DE LA FACTURA EN EL GLOBAL
CAD = "update facventa set total = (select sum(importe) from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'  AND ( CANTIDAD > 0 OR CANTIDADP > 0 ) )  where numfactura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
cn.Execute CAD
'Unload Me
  frmFactDet.AdoFacDet.Refresh

End Sub

Private Sub cmbprod_GotFocus()
RESP = SendMessageLong(cmbprod.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbprod_LostFocus()
On Error GoTo Error:
pr = InStr(cmbprod.Text, "[")
pr1 = InStr(cmbprod.Text, "]")
strcveprod = Trim(Mid(cmbprod.Text, pr + 1, 10))
pr = InStr(cmbprod.Text, "*")
tasas(1) = Mid(cmbprod.Text, pr + 1, 1)
Select Case tasas(1)
    Case 1
        ivas(1) = 0
        iepss(1) = 0
    Case 2
        ivas(1) = 15
        iepss(1) = 0
    Case 3
        ivas(1) = 15
        iepss(1) = 25
    Case 4
        ivas(1) = 15
        iepss(1) = 30
End Select
RESP = SendMessageLong(cmbprod.hwnd, &H14F, False, 1)
txtcajas.SetFocus
Exit Sub
Error:
  Exit Sub
End Sub

Private Sub cmdRegresa_Click()

Hide
'Unload Me
End Sub

Private Sub Command2_Click()
    'SE AGREGA PERO SOLO UN PRODUCTO SIN IVA Y SIN IEPS
    prod = Trim(strcveprod)
    'SERIE = Mid(Factura.Caption, 1, 2)
    'If Mid(SERIE, 2, 1) = "-" Then
    '    SERIE = Mid(SERIE, 1, 1)
    'End If
    'If Len(SERIE) = 1 Then
    '    Factura = Mid(Factura.Caption, 3, Len(Factura.Caption))
    'Else
    '    Factura = Mid(Factura.Caption, 4, Len(Factura.Caption))
    'End If

    nSepara = InStr(1, Factura.Caption, "-")
    SERIE = Mid(Factura.Caption, 1, nSepara - 1)
    Factura = Trim(Mid(Factura.Caption, nSepara + 1, Len(Factura.Caption)))
    
    'Factura = Mid(txtcampos(2).Text, 4, Len(txtcampos(2).Text))
    RESP = MsgBox("Los Datos que estan en pantalla serviran para agregar el Producto" & vbCrLf & "Deseas Continuar..?", vbYesNo, "NVO PRODUCTO")
    venta = ventas.Caption
    If RESP = vbYes Then
        CAD = "insert into facventa_det (producto,cantidad,cantidadp,precio,preciop,costo,costop,importe,iva,ieps,tasaieps,serie,venta,factura,fecha_det,rfc_det) values(" & _
                "'" & prod & "'," & txtcajas.Text & "," & txtpzas.Text & "," & txtuni.Text & "," & txtuni.Text & "," & txtuni.Text & "," & txtuni.Text & "," & txttot.Text & "," & ivas(1) & "," & iepss(1) & "," & tasas(1) & ",'" & Trim(SERIE) & "'," & venta & ",'" & Trim(Factura) & "','" & date & "','COOF970101111')"
        cn.Execute CAD
    End If
    RESP = vbYes
    If RESP = vbYes Then
        'el iva
        CAD = " update facventa set iva = " & _
                " (select sum(importe - importe / (1 + (iva/100)) ) " & _
                " from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' ) " & _
                " Where numfactura = '" & Trim(Factura) & "' And serie = '" & Trim(SERIE) & "'"
        cn.Execute CAD
        'el ieps
        CAD = "update facventa set ieps = " & _
                " (select sum( IMPORTE - (importe / (1 + (iva/100)) / (1 + (IEPS/100)) " & _
                " + (importe - importe / (1 + (iva/100)))  )  ) " & _
                " from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' )" & _
                " Where numfactura = '" & Trim(Factura) & "' And serie = '" & Trim(SERIE) & "'"
        cn.Execute CAD
    End If
    CAD = "update facventa set total = (select sum(importe) from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "')  where numfactura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
    cn.Execute CAD
    frmFactDet.AdoFacDet.Refresh
End Sub

Private Sub Command3_Click()
prod = frmFactDet.AdoFacDet.Recordset!producto
SERIE = Mid(Factura.Caption, 1, 2)
'If Mid(SERIE, 2, 1) = "-" Then
'   SERIE = Mid(SERIE, 1, 1)
'End If
'If Len(SERIE) = 1 Then
'    Factura = Mid(Factura.Caption, 3, Len(Factura.Caption))
'Else
'    Factura = Mid(Factura.Caption, 4, Len(Factura.Caption))
'End If
nSepara = InStr(1, Factura.Caption, "-")
SERIE = Mid(Factura.Caption, 1, nSepara - 1)
Factura = Trim(Mid(Factura.Caption, nSepara + 1, Len(Factura.Caption)))

CAD = "DELETE FACVENTA_det WHERE FACTURA = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' and producto = '" & Trim(prod) & "'"
cn.Execute CAD
CAD = "update facventa set total = (select sum(importe) from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'  AND ( CANTIDAD > 0 OR CANTIDADP > 0 ) )  where numfactura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
frmFactDet.AdoFacDet.Refresh

'Unload Me
End Sub

Private Sub Command5_Click()
'se carga el catalogo de productos
cmbprod.Clear
Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
If MsgBox("Obtener solamente abarrotes?", vbQuestion + vbYesNo) = vbYes Then
   RsCon.Open "SELECT consec,descripc,contenid,str(paquetes) + ' x ' + ltrim(str(contenid,10,3)) + ' ' + medida as medida,paquetes,tasaieps,fecact  FROM tfproduc Where descripc like '%ABARROTES%' order by descripc ", cn, adOpenDynamic, adLockOptimistic, adCmdText
Else
   RsCon.Open "SELECT consec,descripc,contenid,str(paquetes) + ' x ' + ltrim(str(contenid,10,3)) + ' ' + medida as medida,paquetes,tasaieps,fecact  FROM tfproduc where activo = 1 order by descripc ", cn, adOpenDynamic, adLockOptimistic, adCmdText
End If
If RsCon.EOF Then
   MsgBox "No existen Productos en el catalogo"
Else
   While Not RsCon.EOF
   If Not (IsNull(RsCon!medida)) And Not IsNull(RsCon!PAQUETES) Then
                 cmbprod.AddItem RsCon!descripc + " ( " + RsCon!medida + " ) " _
                  + " [ " + RsCon!CONSEC + " ]*" & RsCon!tasaieps
   End If
   RsCon.MoveNext
   Wend
End If
On Error Resume Next
Me.cmbprod.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
Select Case KeyCode
   Case 13
       keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
        KeyCode = 0
   Case 27
       Hide
   Case 112
        Command2_Click
   Case 113
        btncambia_Click
   Case 114
        Command3_Click
   Case 115
        Hide
   Case 116
        txtcajas.SetFocus
   Case 117
        Me.cmbprod.SetFocus
End Select
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
' Routine:           Form_Load
' Description:       type_description_here
' Created by:        Administrador
' Machine:           MOISES
' Date-Time:         09/01/200212:42:36 PM
' Last modification: last_modification_info_here
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
Private Sub Form_Load()
If Me.cmbprod.ListIndex < 1 Then
   'Command5_Click
   'Command5.Enabled = False
End If
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
' Routine:           txtcajas_GotFocus
' Description:       type_description_here
' Created by:        Administrador
' Machine:           MOISES
' Date-Time:         09/01/200212:29:31 PM
' Last modification: last_modification_info_here
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
Private Sub txtcajas_GotFocus()
    'se ponen por default los valores de precios
    'MsgBox strcveprod
    Dim rsb As ADODB.Recordset
    Set rsb = New ADODB.Recordset
    rsb.Open "select * from preprod where preclave = '" & Trim(strcveprod) & "'", cn, adOpenDynamic, adLockOptimistic
    If Not rsb.EOF Then
        Me.txtuni.Text = rsb!PRECIO2
        'Me.txttot.Text = rsb!precio1
        txtprepza.Text = rsb!precio1
    Else
        txtuni.Text = 1
        txtprepza.Text = 1
        'txttot.Text = 1
    End If
    Set rsb = Nothing
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
' Routine:           txtcajas_LostFocus
' Description:       type_description_here
' Created by:        Administrador
' Machine:           MOISES
' Date-Time:         09/01/200201:30:51 PM
' Last modification: last_modification_info_here
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
Private Sub txtcajas_LostFocus()
txttot.Text = (Val(txtcajas.Text) * Val(txtuni.Text)) + (Val(Me.txtpzas.Text) * Val(txtprepza.Text))
txttot.Refresh
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
' Routine:           txtpzas_LostFocus
' Description:       type_description_here
' Created by:        Administrador
' Machine:           MOISES
' Date-Time:         09/01/200201:35:57 PM
' Last modification: last_modification_info_here
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
Private Sub txtpzas_LostFocus()
txttot.Text = (Val(txtcajas.Text) * Val(txtuni.Text)) + (Val(Me.txtpzas.Text) * Val(txtprepza.Text))
txttot.Refresh
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
' Routine:           txttot_KeyPress
' Description:       type_description_here
' Created by:        Administrador
' Machine:           MOISES
' Date-Time:         09/01/200212:47:13 PM
' Last modification: last_modification_info_here
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
Private Sub txttot_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command2.SetFocus
End Sub

Private Sub txtuni_LostFocus()
  txttot.Text = (Val(txtcajas.Text) * Val(txtuni.Text)) + (Val(Me.txtpzas.Text) * Val(Me.txtprepza.Text))
End Sub
