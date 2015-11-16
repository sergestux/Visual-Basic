VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCodBarrCap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos por codigo de Barra"
   ClientHeight    =   4065
   ClientLeft      =   2595
   ClientTop       =   1275
   ClientWidth     =   5490
   Icon            =   "frmCodBarrCap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3690
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "                             Para salir presione la tecla [Esc]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame fra1 
      ForeColor       =   &H80000002&
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtpiezas 
         Alignment       =   2  'Center
         Height          =   320
         Left            =   1320
         TabIndex        =   2
         Text            =   "0"
         Top             =   1020
         Width           =   855
      End
      Begin VB.TextBox txtcajas 
         Alignment       =   2  'Center
         Height          =   320
         Left            =   240
         TabIndex        =   1
         Text            =   "0"
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtcampo 
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Piezas"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Cajas"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Codigo de barras de la caja"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc AdoTemp 
      Height          =   450
      Left            =   3240
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   794
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
      Caption         =   "AdoTemp"
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
   Begin VB.Label lbletiquetas 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label lbletiquetas 
      Alignment       =   2  'Center
      Caption         =   "Productos escaneados:   XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   5055
   End
End
Attribute VB_Name = "frmCodBarrCap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nProd As Integer    'Numero de productos
Private nCajas As Integer
Private lEnt As Boolean
Private rsttemp As ADODB.Recordset
Private clave As String
Private Sub cmdRegresar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
 Set rsttemp = New ADODB.Recordset
 nProd = 0
 lbletiquetas(1).Caption = "Productos escaneados: " + CStr(nProd)
 AdoTemp.ConnectionString = cCadConex
End Sub

Private Sub txtcajas_GotFocus()
   SendKeys "{HOME}"
   SendKeys "+{END}"
End Sub

Private Sub txtcajas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub txtcampo_KeyPress(KeyAscii As Integer)
Dim rsttemp As ADODB.Recordset
Dim existe As Boolean
On Error GoTo error:
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 And Trim(Me.txtcampo.Text) <> "" Then
  'Busco en el catalogo de productos el codigo
  AdoTemp.RecordSource = "SELECT * FROM tfproduc WHERE barraspza = " & Trim(txtcampo.Text)
  AdoTemp.Refresh
  If AdoTemp.Recordset.BOF And AdoTemp.Recordset.EOF Then
     MsgBox "EL ARTICULO CON EL CODIGO DE BARRAS " & Trim(txtcampo.Text) & " NO EXISTE EN EL CATALOGO DE PRODUCTOS", vbExclamation
     txtcampo.Text = ""
     txtcampo.SetFocus
     lbletiquetas(2).Caption = ""
     Exit Sub
  End If
  lbletiquetas(2).Caption = AdoTemp.Recordset!DESCRIPC & Chr(13) & AdoTemp.Recordset!PAQUETES & " x " & CStr(AdoTemp.Recordset!Contenid) & " " & AdoTemp.Recordset!Medida
  clave = AdoTemp.Recordset!CONSEC
  existe = True: lEnt = True
  txtcajas.SetFocus
 Select Case nOp
 Case 0  'Se Llama desde Pedidos por proveedor al recibir pedido (Entrada)
 
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT descripc, LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT FROM tfproduc WHERE BarrasPza = " & txtcampo.Text, cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rsttemp.RecordCount > 0 Then
       lbletiquetas(2).Caption = rsttemp!DESCRIPC & Chr(13) & rsttemp!Present
    Else
       MsgBox "NO EXISTE PRODUCTO CON EL CODIGO DE BARRAS ESPECIFICADO", vbExclamation
       lbletiquetas(2).Caption = ""
    End If
    txtcampo.Text = ""
    Exit Sub
    txtcajas.Text = 0: txtpiezas.Text = 0
    txtcajas.SetFocus
 Case 1  'Desde BackOrder
 
 Case 2  'Desde Traslado por pedido 'PENDIENTEEEEEEEE SI SE ACEPTAN ENVIAR MAS CAJAS DE LAS SOLICITADAS EN EL PEDIDO
    lEnt = False
    frmTrasladaEnv.AdoDetPed.Recordset.MoveFirst
    frmTrasladaEnv.AdoDetPed.Recordset.Find "DT_PRODUCTO = '" & clave & "'"
    If frmTrasladaEnv.AdoDetPed.Recordset.EOF Then
       existe = False
    Else
       nCajas = frmTrasladaEnv.AdoDetPed.Recordset!df_cantsol
    End If
    Trasl = Trim(frmTrasladaEnv.txtcampos(0).Text)
 Case 3  'Desde envio de traslado ABIERTO, NO EXISTE restriccion para enviar numero de cajas
    lEnt = False
    
 End Select

 If Not existe Then
     MsgBox "LA CLAVE DEL ARTICULO " & clave & " NO EXISTE EN EL PEDIDO SELECCIONADO ", vbExclamation
     txtcampo.Text = ""
     txtcampo.SetFocus
     Exit Sub
 End If
 
 'Valido que no reciban mas cajas de las solicitadas
 If nOp = 0 Then      'Recibiendo pedido por proveedor
    If nCajas > Val(frmPedProv.AdoDetGlo.Recordset!dg_cantreal) Then
       frmPedProv.AdoDetGlo.Recordset!dg_cantreal = frmPedProv.AdoDetGlo.Recordset!dg_cantreal + 1
       frmPedProv.AdoDetGlo.Recordset.Update
    Else
       MsgBox "NO PUEDE RECIBIR MAS CAJAS DE LAS SOLICITADAS", vbCritical
       txtcampo.Text = ""
       txtcampo.SetFocus
       Exit Sub
    End If
 ElseIf nOp = 1 Then  'Recibiendo BackOrder
 
       MsgBox "NO PUEDE RECIBIR MAS CAJAS DE LAS SOLICITADAS", vbCritical
       txtcampo.Text = ""
       txtcampo.SetFocus
       Exit Sub
    'End If
 ElseIf nOp = 2 Then  'Enviando traslado cerrado por pedido
    If nCajas > Val(frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidad) Then
       frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidad = frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidad + 1
       frmTrasladaEnv.AdoDetPed.Recordset.Update
    Else
       MsgBox "NO PUEDE ENVIAR MAS CAJAS DE LAS SOLICITADAS POR LA TIENDA", vbCritical
       txtcampo.Text = ""
       txtcampo.SetFocus
       Exit Sub
    End If
 ElseIf nOp = 3 Then  'Enviando Traslado Abierto. Pueden enviar lo que quieran
 End If
 
 nProd = nProd + 1
 
 lbletiquetas(1).Caption = "Productos escaneados: " + CStr(nProd)
 End If
 Exit Sub
error:
  MsgBox Err.Description
End Sub


Private Sub txtpiezas_GotFocus()
   SendKeys "{HOME}"
   SendKeys "+{END}"
End Sub

Private Sub txtpiezas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub txtpiezas_LostFocus()
Select Case nOp
Case 3
    If frmTrasladaEnv.AdoTraslada.Recordset!t_enviado Then Exit Sub
    If MsgBox("CONFIRMA SI DESEAS AGREGAR/MODIFICAR EL PRODUCTO SELECCIONADO", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If Not (frmTrasladaEnv.AdoDetPed.Recordset.BOF And frmTrasladaEnv.AdoDetPed.Recordset.BOF) Then frmTrasladaEnv.AdoDetPed.Recordset.MoveFirst
    frmTrasladaEnv.AdoDetPed.Recordset.Find "dt_producto = '" & clave & "'"
    If frmTrasladaEnv.AdoDetPed.Recordset.EOF Then
        frmTrasladaEnv.AdoDetPed.Recordset.AddNew
    End If
    frmTrasladaEnv.AdoDetPed.Recordset!DT_CLAVE = frmTrasladaEnv.txtcampos(0).Text
    frmTrasladaEnv.AdoDetPed.Recordset!Dt_producto = clave
    frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidad = IIf(IsNull(frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidad), 0, frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidad) + txtcajas.Text
    frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidadp = IIf(IsNull(frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidadp), 0, frmTrasladaEnv.AdoDetPed.Recordset!dt_cantidadp) + txtpiezas.Text
    frmTrasladaEnv.AdoDetPed.Recordset.Update
    frmTrasladaEnv.AdoDetPed.Refresh
    frmTrasladaEnv.AdoDetPed.Recordset.MoveFirst
    frmTrasladaEnv.AdoDetPed.Recordset.Find "DT_PRODUCTO = '" & clave & "'"
    txtcampo.Text = ""
    txtcampo.SetFocus
End Select
End Sub
