VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTrasladaRec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibir Traslado"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   Icon            =   "frmRecTraslada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "t_costo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "AdoTraslada"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "T_fecharec"
      DataSource      =   "AdoTraslada"
      Height          =   285
      Index           =   1
      Left            =   8160
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc AdoTraslada 
      Height          =   330
      Left            =   4560
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "AdoTraslada"
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
   Begin VB.CheckBox chkRecibido 
      Caption         =   "Traslado recibido"
      DataField       =   "t_recibido"
      DataSource      =   "AdoTraslada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox PicBotones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   11490
      TabIndex        =   3
      Top             =   7440
      Width           =   11490
      Begin VB.PictureBox Rpt 
         Height          =   480
         Left            =   960
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   13
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdboton 
         Caption         =   "&Reporte"
         Height          =   495
         Index           =   2
         Left            =   3480
         Picture         =   "frmRecTraslada.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdboton 
         Caption         =   "&Grabar"
         Height          =   495
         Index           =   0
         Left            =   5400
         Picture         =   "frmRecTraslada.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdboton 
         Caption         =   "&Cancelar"
         Height          =   495
         Index           =   1
         Left            =   7320
         Picture         =   "frmRecTraslada.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc AdoDetTra 
      Height          =   330
      Left            =   960
      Top             =   6600
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
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
      ConnectStringType=   3
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
      Caption         =   "AdoTrasl"
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
   Begin MSDataGridLib.DataGrid dbgrdDetTrasl 
      Bindings        =   "frmRecTraslada.frx":0D20
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1.5
      RowHeight       =   15
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
      Caption         =   "DESGLOSE DEL TRASLADO"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "dt_pedido"
         Caption         =   "        PEDIDO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "dt_producto"
         Caption         =   "     CLAVE PROD."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Descripc"
         Caption         =   "        DESCRIPCION DEL PRODUCTO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Medida"
         Caption         =   "        MEDIDA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "df_cantsol"
         Caption         =   "SOL CAJAS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "dt_Cantidad"
         Caption         =   "REC. CAJAS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "df_cantsolp"
         Caption         =   "SOL. PZA."
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
      BeginProperty Column07 
         DataField       =   "dt_cantidadp"
         Caption         =   "REC. PZA."
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
            Alignment       =   1
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3209.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCampos 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin ComctlLib.StatusBar stbmensajes 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   8175
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "                                                                                              Para salir presione la tecla [Esc]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Monto del traslado"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Fecha de recepcion"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Folio del traslado"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTrasladaRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkRecibido_Click()
    lbletiquetas(1).Visible = chkRecibido.Value = 1
    txtcampos(1).Visible = chkRecibido.Value = 1
    If AdoTraslada.Recordset!T_RECIBIDO = 0 Then txtcampos(1).Text = date + Time
    txtcampos(1).Enabled = False
End Sub

Private Sub cmdBoton_Click(Index As Integer)
Dim rstInv As ADODB.Recordset
Dim rstProPed As ADODB.Recordset
Dim rstPedAnt
Dim lTrans As Boolean
On Error GoTo Error
lTrans = False
Select Case Index
Case 0  'Grabrar recibir traslado
     If chkRecibido.Value = 0 Then
         MsgBox "ES NECESARIO ACTIVAR LA CASILLA DE TRASLADO RECIBIDO", vbExclamation
         chkRecibido.SetFocus
         Exit Sub
     End If
     RESP = MsgBox("DESEAS RECIBIR EL TRASLADO Y AFECTAR INVENTARIO", vbQuestion + vbYesNo)
     If RESP = vbYes Then
        cn.BeginTrans: lTrans = True
        If cn.State = 0 Then cn.Open
        'Cargo el inventario para actualizar existencias
        Set rstInv = New ADODB.Recordset
        rstInv.ActiveConnection = cn
        rstInv.CursorType = adOpenKeyset
        rstInv.LockType = adLockOptimistic
        rstInv.Source = "SELECT * FROM [DetalleTraslado],[Inventario] WHERE DetalleTraslado.dt_producto *= Inventario.inprod " & " AND DetalleTraslado.dt_clave ='" & txtcampos(0).Text & "'"
        rstInv.Open
           
        Set rstProPed = New ADODB.Recordset
        rstProPed.Open "SELECT * FROM Tfproduc,DetalleTraslado WHERE DETALLETRASLADO.dt_producto = TFPRODUC.Consec AND dt_clave = '" & txtcampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not (AdoDetTra.Recordset.BOF And AdoDetTra.Recordset.EOF) Then AdoDetTra.Recordset.MoveFirst
        rstPedAnt = ""
        While Not AdoDetTra.Recordset.EOF
            StbMensajes.SimpleText = Space(50) + "Actualizando " + LCase(AdoDetTra.Recordset!descripc + " " + AdoDetTra.Recordset!medida) + " del pedido " + IIf(IsNull(AdoDetTra.Recordset!dt_pedido), AdoDetTra.Recordset!Dt_producto, AdoDetTra.Recordset!dt_pedido)
            StbMensajes.Refresh
            'Primero hago la busqueda en el catalogo de articulos para saber el numero de piezas de la caja
            rstProPed.MoveFirst
            rstProPed.Find "dt_producto = '" & AdoDetTra.Recordset!Dt_producto & "'"
            If rstProPed.EOF Then
                MsgBox "EL ARTICULO CON CODIGO " & AdoDetTra.Recordset!Dt_producto & "NO EXISTE EN EL CATALOGO DE PRODUCTOS" & _
                "A CONTINUACION SE DESAHARAN LOS CAMBIOS REALIZADOS", vbCritical
                cn.RollbackTrans
                Exit Sub
            End If
            rstInv.MoveFirst
            rstInv.Find "inProd = '" & AdoDetTra.Recordset!Dt_producto & "'"
            If rstInv.EOF Then
               MsgBox "LA CLAVE DEL PRODUCTO " + AdoDetTra.Recordset!Dt_producto + " NO EXISTE EN EL INVENTARIO"
            Else
                rstInv!InCant = rstInv!InCant + (AdoDetTra.Recordset!dt_cantidad * rstProPed!PAQUETES) + AdoDetTra.Recordset!dt_cantidadp
                rstInv.Update
            End If
            If AdoTraslada.Recordset!t_tipo = 0 Then
                If rstPedAnt <> AdoDetTra.Recordset!dt_pedido Then
                    rstPedAnt = AdoDetTra.Recordset!dt_pedido
                    cn.Execute "UPDATE Pedidos set p_recibido = 1, p_fecEntReal = '" & txtcampos(1).Text & "' WHERE p_pedido = '" & rstPedAnt & "'"
                End If
                cn.Execute "UPDATE DetalleFactura set df_cantreal = " & AdoDetTra.Recordset!dt_cantidad & " WHERE df_pedido = '" & rstPedAnt & "' AND df_prod = '" & AdoDetTra.Recordset!Dt_producto & "'"
            End If
            AdoDetTra.Recordset.MoveNext
        Wend
        AdoTraslada.Recordset!t_fecharec = date + Time
        AdoTraslada.Recordset.Update
        cn.CommitTrans
        Set rstInv = Nothing
        Unload Me
     End If
Case 2
     cMensaje = StbMensajes.SimpleText
    StbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
    StbMensajes.Refresh
    Rpt.Connect = cCadConex
    Rpt.ReportFileName = App.Path & IIf(AdoTraslada.Recordset!t_tipo = False, "\TraslRec.rpt", "\TrasAbto.rpt")
    Rpt.WindowTitle = "Recepcion del traslado con folio " & txtcampos(0).Text
    Rpt.Formulas(0) = "FORMSELEC = '" & txtcampos(0).Text & "'"
    Rpt.Formulas(1) = "TRASLADO= 'RECEPCION DEL TRASLADO CON FOLIO " & txtcampos(0).Text & "'"
    Rpt.Formulas(2) = "CANTRECTRA= 'CAJA. REC'"
    If AdoTraslada.Recordset!t_tipo = False Then Rpt.Formulas(3) = "CANTRECPZA= 'PZA. REC'"
    Rpt.Action = 1
    StbMensajes.SimpleText = cMensaje
    StbMensajes.Refresh

Case 1  'Regresar a la pantalla principal
     Unload Me
End Select
Exit Sub
Error:
   MsgBox "OCURRIO EL SIGUIENTE ERROR: " + Chr(13) + UCase(Err.Description), vbCritical
   If lTrans Then
      MsgBox "A CONTINUACION SE DESHARAN LAS MODIFICACIONES REALIZADAS AL INVENTARIO", vbCritical
      cn.RollbackTrans
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmtraslados.Show
End Sub

Private Sub txtCampos_GotFocus(Index As Integer)
  FrmReport.Hide
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0
    txtcampos(0).Text = Trim(txtcampos(0).Text)
    txtcampos(0).Refresh
    If txtcampos(0).Text = "" Or IsNull(txtcampos(0)) Then
        MsgBox "No puede dejar en blanco el folio del traslado"
        txtcampos(0).SetFocus
        Exit Sub
    End If
    
    AdoTraslada.LockType = adLockOptimistic
    AdoTraslada.ConnectionString = cCadConex
    AdoTraslada.CursorType = adOpenKeyset
    AdoTraslada.RecordSource = "SELECT * FROM TRASLADOS WHERE t_clave = '" & txtcampos(0).Text & "'"
    AdoTraslada.Refresh
    If AdoTraslada.Recordset.BOF And AdoTraslada.Recordset.EOF Then
        MsgBox "No existe el folio especificado", vbExclamation
        txtcampos(0).SetFocus
        Exit Sub
    End If
    txtcampos(0).Enabled = False
    AdoDetTra.ConnectionString = cn.ConnectionString
    AdoDetTra.CommandType = adCmdText
    'Si es traslado cerrado = pedido
    If AdoTraslada.Recordset!t_tipo = False Then
        AdoDetTra.RecordSource = "SELECT DetalleTraslado.dt_pedido,DetalleTraslado.dt_producto,DetalleTraslado.dt_producto, TfProduc.descripc, LTrim(str(paquetes)) + ' X ' + LTRIM( str(contenid)) + ' ' + MEDIDA as MEDIDA,DetalleTraslado.dt_cantidad,DetalleTraslado.dt_cantidadp, DetalleFactura.df_prod,DetalleFactura.df_cantsol,DetalleFactura.df_cantsolp FROM DetalleTraslado,tfproduc," & _
                                 "DetalleFactura WHERE DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' AND DetalleTraslado.dt_producto = Tfproduc.Consec AND DetalleTraslado.dt_producto = Detallefactura.df_prod AND DetalleTraslado.dt_pedido = DetalleFactura.df_pedido ORDER BY dt_pedido,dt_producto"
    Else
        AdoDetTra.RecordSource = "SELECT DetalleTraslado.dt_pedido,DetalleTraslado.dt_producto,DetalleTraslado.dt_producto, TfProduc.descripc, LTrim(str(paquetes)) + ' X ' + LTRIM( str(contenid)) + ' ' + MEDIDA as MEDIDA,DetalleTraslado.dt_cantidad, DetalleTraslado.dt_cantidadP FROM DetalleTraslado,tfproduc " & _
                                 " WHERE DetalleTraslado.dt_clave = '" & txtcampos(0).Text & "' AND DetalleTraslado.dt_producto = Tfproduc.Consec ORDER BY dt_producto"
    End If
    AdoDetTra.Refresh
    
    Me.dbgrdDetTrasl.Visible = True
    cmdboton(0).Visible = True
    cmdboton(1).Visible = True
    cmdboton(2).Visible = True
    chkRecibido.Visible = True
    lbletiquetas(2).Visible = True
    txtcampos(2).Visible = True
    txtcampos(2).Enabled = False
    cmdboton(0).Enabled = chkRecibido.Value = 0
    chkRecibido.Enabled = chkRecibido.Value = 0
    
    Set rstTra = Nothing
End Select
Exit Sub
Error:
MsgBox Err.Description
End Sub
