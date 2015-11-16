VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmBackDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle backorder"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmBackDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cRpt 
      Left            =   120
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowLeft      =   0
      WindowTop       =   0
      WindowState     =   2
   End
   Begin MSComctlLib.StatusBar stbMensajes 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   8250
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "                                                                                           Para salir presione la tecla  [ESC]"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dggrdDetBack 
      Bindings        =   "frmBackDet.frx":0442
      Height          =   6735
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11880
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1.5
      RowHeight       =   15
      TabAction       =   2
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "producto"
         Caption         =   "CLAVE"
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
         DataField       =   "Descripc"
         Caption         =   "DESCRIPCION"
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
         DataField       =   "medida"
         Caption         =   "MEDIDA"
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
         DataField       =   "cantasurtir"
         Caption         =   "CAJAS PEND."
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
         DataField       =   "cantasurtirp"
         Caption         =   "PZAS. PEND"
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
      BeginProperty Column05 
         DataField       =   "cantrecibida"
         Caption         =   "CAJ.REC."
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
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   5249.764
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   975.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoDetBack 
      Height          =   330
      Left            =   840
      Top             =   7200
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "AdoDetBack"
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
   Begin VB.PictureBox PicBotones 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   11850
      TabIndex        =   6
      Top             =   7620
      Width           =   11910
      Begin VB.CommandButton cmdCodBarra 
         Caption         =   "Cod. &Barra"
         Height          =   400
         Left            =   4560
         Picture         =   "frmBackDet.frx":045B
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Entrada al inventario por medio de etiquetas"
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdboton 
         Caption         =   "R&eporte"
         Height          =   400
         Index           =   2
         Left            =   3000
         Picture         =   "frmBackDet.frx":0591
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Vista preliminar de las entregas realizadas por pedido"
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdboton 
         Caption         =   "&Regresar"
         Height          =   400
         Index           =   1
         Left            =   7680
         Picture         =   "frmBackDet.frx":0AC3
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Regresar a la pantalla de pedidos por proveedor"
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdboton 
         Caption         =   "&Grabar"
         Height          =   400
         Index           =   0
         Left            =   6120
         Picture         =   "frmBackDet.frx":0C35
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Grabar backorder y aumentar inventario"
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.Label lblEnca 
      Alignment       =   2  'Center
      Caption         =   "PROVEEDOR : XXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   11535
   End
   Begin VB.Label lblEnca 
      Alignment       =   2  'Center
      Caption         =   "PRODUCTOS PENDIENTES  DE ENTREGAR DEL PEDIDO  XX"
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
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   11295
   End
End
Attribute VB_Name = "frmBackDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lExiste As Boolean
Private Sub cmdBoton_Click(Index As Integer)
Dim rsttemp As ADODB.Recordset
Dim rstDetPed As ADODB.Recordset
Dim nSutido As Integer
On Error GoTo Error
Select Case Index
Case 0 'Grabar
      'Cargo el detalle del pedido global solo aquellos que difiere la cantidad solicitada con la recibida
      'Para incrementarle el backorder recibido
      'Set rstDetPed = New ADODB.Recordset
      'rstDetPed.Open "SELECT * FROM DETALLEGLOBAL WHERE dg_pedido = '" & frmpedBod.dbgrdPed.Columns(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
      'Cargo el inventario solo de aquellos del DetalleBackorder
      Set rsttemp = New ADODB.Recordset
      rsttemp.ActiveConnection = cn
      'rsttemp.Open "INVENTARIO", cn, adOpenKeyset, adLockOptimistic, adCmdTable
      AdoDetBack.Refresh
      AdoDetBack.Recordset.MoveFirst
      While Not AdoDetBack.Recordset.EOF
           StbMensajes.SimpleText = Space(75) & "Actualizando producto " & AdoDetBack.Recordset!producto
           StbMensajes.Refresh
           'Si es diferente el faltante con la cantidad recibida
           nsurtido = 0
           If (AdoDetBack.Recordset!cantasurtir <> AdoDetBack.Recordset!cantrecibida) Or (AdoDetBack.Recordset!CantAsurtirP <> AdoDetBack.Recordset!CantRecibidaP) Then
              'Cargo nuevamente el faltante al DetalleBack
              If AdoDetBack.Recordset!cantrecibida <> 0 Then
                 'Obtiene la cantidad en piezas a incrementar en el inventario
                 cn.Execute "UPDATE DETALLEBACK SET Situacion = 0, FECHA = getdate() WHERE producto = '" & AdoDetBack.Recordset!producto & "' AND pedidog = '" & frmpedBod.dbgrdPed.Columns(0).Text & "' AND SITUACION = 1"
                 nsurtido = AdoDetBack.Recordset!cantrecibida
                 cn.Execute "INSERT INTO [DetalleBack](NoBack,producto,cantAsurtir,CantRecibida,cantasurtirP,cantRecibidaP,pedidog,fecha,situacion) VALUES " & _
                           "('" & AdoDetBack.Recordset!NoBack & "','" & AdoDetBack.Recordset!producto & "','" & AdoDetBack.Recordset!cantasurtir - AdoDetBack.Recordset!cantrecibida & "','0','" & AdoDetBack.Recordset!CantAsurtirP - AdoDetBack.Recordset!CantRecibidaP & "','0','" _
                           & frmpedBod.dbgrdPed.Columns(0).Text & "','" & Date & "','1' )"
              End If
           'Se surtio todo el faltante
           Else
              nsurtido = AdoDetBack.Recordset!cantrecibida
              cn.Execute "UPDATE DETALLEBACK SET situacion = 0, fecha = getdate() WHERE producto = '" & AdoDetBack.Recordset!producto & "' AND pedidog = '" & frmpedBod.dbgrdPed.Columns(0).Text & "' AND SITUACiON = 1"
           End If
           If nsurtido > 0 Then
              rsttemp.Open "SELECT * FROM inventario WHERE Inprod = '" & AdoDetBack.Recordset!producto & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
              If rsttemp.BOF And rsttemp.EOF Then
                 MsgBox "NO EXISTE EN EL INVENTARIO EL PRODUCTO CON LA CLAVE " & AdoDetBack.Recordset!producto & Chr(13) & " A CONTINUACION SE DARA DE ALTA", vbExclamation
                 cn.Execute "INSERT INTO [Inventario] (inprod,insucursal,incant,ininicial,instock) VALUES " _
                            & "('" & AdoDetBack.Recordset!producto & "','" & Mid(cSucursal, 1, 3) & "','" & nsurtido & "','0','0')"
              Else
                 rsttemp!InCant = rsttemp!InCant + nsurtido
                 rsttemp.Update
              End If
              rsttemp.Close
           End If
           AdoDetBack.Recordset.MoveNext
      Wend
      Unload Me
Case 1 'Regresar
     Unload Me
Case 2 'Reporte
     cMensaje = StbMensajes.SimpleText
     StbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
     StbMensajes.Refresh
     cRpt.Connect = cCadConex
     cRpt.ReportFileName = App.Path & "\PrBackor.rpt"
     cRpt.WindowTitle = "Entregas del pedido " & frmpedBod.dbgrdPed.Columns(0).Text
     cRpt.Formulas(0) = "FORMSELEC = '" & frmpedBod.dbgrdPed.Columns(0).Text & "'"
     cRpt.Formulas(1) = "PEDIDO = 'BACKORDER DEL PEDIDO " & frmpedBod.dbgrdPed.Columns(0).Text & "'"
     cRpt.Formulas(2) = "PROVED = 'PROVEEDOR " & frmpedBod.dbgrdPed.Columns(1).Text & Space(2) & frmpedBod.cmbproved.Text & "'"
     cRpt.Action = 1
     StbMensajes.SimpleText = cMensaje
     StbMensajes.Refresh
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdCodBarra_Click()
  nOp = 1  'Para que la forma de lectura de codigos de barra sepa de donde se esta llamando
  frmCodBarrCap.Show 1
End Sub

Private Sub dbgrdRec_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next
If ColIndex = 0 And AdoDetBack.Recordset!cantrecibida > AdoDetBack.Recordset!cantasurtir Then
   MsgBox "LA CANTIDAD RECIBIDA EN CAJAS NO PUEDE SER MAYOR AL FALTANTE", vbExclamation
   AdoDetBack.Recordset!cantrecibida = 0
ElseIf ColIndex = 1 And AdoDetBack.Recordset!CantRecibidaP > AdoDetBack.Recordset!CantAsurtirP Then
   MsgBox "LA CANTIDAD RECIBIDA EN PIEZAS NO PUEDE SER MAYOR AL FALTANTE", vbExclamation
   AdoDetBack.Recordset!CantRecibidaP = 0
End If
End Sub

Private Sub dbgrdRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub dbgrdRec_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Error:
If Not (AdoDetBack.Recordset.BOF And AdoDetBack.Recordset.EOF) Then AdoDetBack.Recordset.MoveFirst
AdoDetBack.Recordset.Find "CLAVE = '" & AdoDetBack.Recordset!producto & "'"
Error:
End Sub

Private Sub dggrdDetBack_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
 'ARUAL
 cn.Execute "UPDATE DETALLEBACK SET cantrecibida = " & Me.dggrdDetBack.Columns(ColIndex).Text & " WHERE PRODUCTO = '" & AdoDetBack.Recordset!producto & "' AND Pedidog = '" & Trim(Me.AdoDetBack.Recordset!pedidog) & "' AND situacion = 1"
 Cancel = True
 AdoDetBack.Refresh
End Sub

Private Sub Form_Activate()
'Si no existen pedidos por tienda de este pedido descargo la forma
'If Not lExiste Then Unload Me
End Sub

Private Sub Form_Load()
Dim lvalor As String
Dim rsttemp As ADODB.Recordset
 'Cargo el 2do. Grid con la cantidad recibida
On Error GoTo Error:
 
 lvalor = "0"
 
 'AdoDetBack.ConnectionString = cCadConex
 'AdoDetBack.CommandType = adCmdText
 'AdoDetBack.RecordSource = "SELECT * FROM DetalleBack WHERE pedidog = '" & frmpedBod.dbgrdPed.Columns(0).Text & "' AND situacion = 1 ORDER BY producto"
 'AdoDetBack.Refresh
 'Si no hay Backorder activo o no existe
 'If AdoDetBack.Recordset.BOF And AdoDetBack.Recordset.EOF Then
 '   'Busco si existen inactivos
 '   AdoDetBack.RecordSource = "SELECT * FROM DetalleBack WHERE pedidog = '" & frmpedBod.dbgrdPed.Columns(0).Text & "' AND situacion = 0 ORDER BY producto"
 '   AdoDetBack.Refresh
 '   lExiste = True
 '   If AdoDetBack.Recordset.BOF And AdoDetBack.Recordset.BOF Then
 '      MsgBox "NO EXISTE BACKORDER PARA ESTE PEDIDO", vbExclamation
 '      lExiste = False
 '   Else
 '      lvalor = "0"
 '   End If
 'End If
 Set rsttemp = New ADODB.Recordset
 rsttemp.Open "SELECT * FROM DetalleBack WHERE pedidog = '" & frmpedBod.dbgrdPed.Columns(0).Text & "' AND situacion = 1 ORDER BY producto", cn, adOpenKeyset, adLockOptimistic, adCmdText
 If rsttemp.RecordCount > 0 Then
    lvalor = "1"
 End If
 'Cargo el primer grid con los datos de articulos por entregar activos
 AdoDetBack.ConnectionString = cCadConex
 AdoDetBack.CommandType = adCmdText
 AdoDetBack.RecordSource = "SELECT  Noback, situacion, producto, descripc, LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, cantasurtir, cantasurtirP, cantrecibida, CantrecibidaP,pedidog FROM DetalleBack, tfproduc WHERE DetalleBack.producto = TFPRODUC.consec AND pedidog = '" & frmpedBod.dbgrdPed.Columns(0).Text & "' AND situacion = " & lvalor & " ORDER BY descripc, contenid"
 AdoDetBack.Refresh
 'Me.dggrdDetBack.Splits(0).Locked = True
 
 If lvalor = "1" Then
    lblEnca(1).ForeColor = &H8000000D
    lblEnca(1).Caption = "PRODUCTOS FALTANTES DE ENTREGAR EN EL PEDIDO " & frmpedBod.dbgrdPed.Columns(0).Text
 Else
    lblEnca(1).Caption = "PRODUCTOS QUE YA SE SURTIERON EN EL PEDIDO " & frmpedBod.dbgrdPed.Columns(0).Text
    cmdboton(0).Enabled = False
    cmdCodBarra.Enabled = False
    'dbgrdRec.Splits(0).Locked = True
 End If
 lblEnca(0).Caption = "PROVEEDOR " & frmpedBod.dbgrdPed.Columns(1).Text & Space(2) & frmpedBod.cmbproved.Text
Exit Sub
Error:
    MsgBox Err.Description
End Sub
