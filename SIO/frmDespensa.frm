VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmdespensa 
   Caption         =   "Módulo de Despensas"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   Icon            =   "frmDespensa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   450
      Index           =   0
      Left            =   120
      Picture         =   "frmDespensa.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Ticket"
      Enabled         =   0   'False
      Height          =   450
      Index           =   2
      Left            =   1800
      Picture         =   "frmDespensa.frx":05B4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdOpcion 
      Appearance      =   0  'Flat
      Caption         =   "&Regresar"
      Enabled         =   0   'False
      Height          =   450
      Index           =   1
      Left            =   960
      Picture         =   "frmDespensa.frx":0726
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid dtgd 
      Bindings        =   "frmDespensa.frx":0898
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      BorderStyle     =   0
      HeadLines       =   1.5
      RowHeight       =   15
      RowDividerStyle =   3
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Caption         =   "PRODUCTOS QUE INCLUYE LA DESPENSA"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "clave"
         Caption         =   "Clave"
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
      BeginProperty Column01 
         DataField       =   "cantidad"
         Caption         =   "Cantidad"
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
      BeginProperty Column02 
         DataField       =   "descripc"
         Caption         =   "Descripción"
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
      BeginProperty Column03 
         DataField       =   "contenid"
         Caption         =   "Contenido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "medida"
         Caption         =   "Medida"
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
            Alignment       =   2
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   3449.764
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dtga 
      Bindings        =   "frmDespensa.frx":08AE
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12632256
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1.5
      RowHeight       =   17
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "consec"
         Caption         =   "Clave"
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
      BeginProperty Column01 
         DataField       =   "descripc"
         Caption         =   "Descripción"
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
      BeginProperty Column02 
         DataField       =   "contenid"
         Caption         =   "Contenido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "medida"
         Caption         =   "Medida"
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
      BeginProperty Column04 
         DataField       =   "precio"
         Caption         =   "Precio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3449.764
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   3120
      Top             =   4800
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "AdoDesp"
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
   Begin MSAdodcLib.Adodc AdoDesp 
      Height          =   330
      Left            =   360
      Top             =   4800
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "AdoDesp"
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
   Begin VB.TextBox txtNumprod 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtNumDesp 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lbletiqueta 
      Alignment       =   2  'Center
      Caption         =   "A FAVOR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblDispo 
      Alignment       =   1  'Right Justify
      Caption         =   "$ 0.00"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label lbletiqueta 
      Alignment       =   2  'Center
      Caption         =   "Nº de Productos:"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lbletiqueta 
      Alignment       =   2  'Center
      Caption         =   "Nº de Vales"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu MnuDesp 
      Caption         =   "&Despensa"
      Visible         =   0   'False
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
      End
   End
End
Attribute VB_Name = "frmdespensa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nDesp As Currency
Private Tabla As String

Private Sub totDesp()
 On Error Resume Next
 Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
 rs.Open "SELECT SUM(IMPORTE) AS impte FROM " & Tabla, cn, adOpenForwardOnly, adLockOptimistic, adcmdtex
 lblDispo.Caption = Format(nDesp - IIf(IsNull(rs!impte), 0, rs!impte), "$ ####,###,#00.00")
 Me.dtga.Visible = rs!impte > 0
 Me.lbletiqueta(2).Visible = Val(Format(nDesp - IIf(IsNull(rs!impte), 0, rs!impte), "########00.00")) > 0
 rs.Close
 Set RST = Nothing
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0
     For N = 1 To txtNumDesp.Text
        Folio = InputBox("INTRODUZCA FOLIO DEL VALE " & N, "Folio del Vale")
        While Not IsNumeric(Folio)
            Folio = InputBox("INTRODUZCA FOLIO DEL VALE", "Folio del Vale")
        Wend
        If Me.dtga.Visible = True Then  'Cuando cambian productos de la despensa
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
            AdoDesp.Refresh
            rs.Open "SELECT * FROM despensa WHERE despensa = 1", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not rs.EOF
               Me.AdoDesp.Recordset.MoveFirst
               AdoDesp.Recordset.Find "clave = '" & Trim(rs!clave) & "'"
               'Los borrados de la despensa
               If AdoDesp.Recordset.EOF Then
                  cn.Execute "INSERT INTO despdet(folio,clave,cantidad) VALUES (" & Folio & ",'" & rs!clave & "'," & rs!cantidad * -1 & ")"
               Else
                  'Lo capturado es MAYOR a lo que trae la despensa
                  If AdoDesp.Recordset!cantidad > rs!cantidad Then
                     cn.Execute "INSERT INTO despdet(folio,clave,cantidad) VALUES (" & Folio & ",'" & rs!clave & "'," & AdoDesp.Recordset!cantidad - rs!cantidad & ")"
                  'Lo capturado es MENOR a lo que trae la despensa
                  ElseIf AdoDesp.Recordset!cantidad < rs!cantidad Then
                     cn.Execute "INSERT INTO despdet(folio,clave,cantidad) VALUES (" & Folio & ",'" & rs!clave & "'," & AdoDesp.Recordset!cantidad - rs!cantidad & ")"
                  'Se cambia por otro producto diferente al que trae la despensa
                  'Else
                  '   cn.Execute "INSERT INTO despdet(folio,clave,cantidad) VALUES (" & Folio & ",'" & rs!clave & "'," & AdoDesp.Recordset!cantidad - rs!cantidad & ")"
                  End If
                  AdoDesp.Recordset!GRABADO = 1
                  AdoDesp.Recordset.Update
               End If
               rs.MoveNext
            Wend
            If Val(Format(lblDispo.Caption, "###,###,##0.00")) <> 0 Then
               cn.Execute "INSERT INTO despdet(folio,clave,cantidad) VALUES (" & Folio & ",'1008833'," & -1 * Val(Format(lblDispo.Caption, "###,###,##0.00")) & ")"
            End If
            cn.Execute "INSERT INTO despdet(folio,clave,cantidad) SELECT fol= " & Folio & ",clave,cantidad FROM " & Tabla & " WHERE grabado = 0"
        Else
            cn.Execute "INSERT INTO despdet(folio,clave,cantidad) VALUES (" & Folio & ",'3000336',1)"
        End If
        cn.Execute "DELETE FROM " & Tabla
        Me.AdoDesp.Refresh
        Me.txtNumDesp.Enabled = True
        Me.cmdOpcion(0).Enabled = False
        Me.cmdOpcion(1).Enabled = False
        Me.cmdOpcion(2).Enabled = False
        Me.dtga.Visible = False
        Me.lbletiqueta(2).Visible = False
        Me.lblDispo.Caption = "$0.00"
        MsgBox "LA INFORMACION SE GRABO CORRECTAMENTE", vbInformation, "Despensas"
        
     Next
Case 1
       Unload Me
Case 2
     ticdesp   'Ticket de despensa
End Select
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub dtga_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Me.dtga.Visible = False
If KeyAscii = 13 Then
   If MsgBox("DESEAS AGREGAR EL PRODUCTO " & Me.AdoArt.Recordset!descripc & " " & AdoArt.Recordset!CONTENID & " " & AdoArt.Recordset!medida, vbYesNo + vbQuestion) = vbYes Then
      cn.Execute "INSERT INTO " & Tabla & " (clave,descripc,contenid,medida,cantidad,precio,importe) VALUES ('" & AdoArt.Recordset!CONSEC & "','" & AdoArt.Recordset!descripc & "'," & AdoArt.Recordset!CONTENID & ",'" & AdoArt.Recordset!medida & "',1," & AdoArt.Recordset!PRECIO & "," & AdoArt.Recordset!PRECIO & ")"
   End If
   AdoDesp.Refresh
   totDesp
End If
End Sub

Private Sub dtgd_AfterUpdate()
  totDesp
End Sub

Private Sub dtgd_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo Error:
AdoDesp.Recordset!importe = AdoDesp.Recordset!PRECIO * Me.dtgd.Columns(1).Text
If ColIndex = 1 Then SendKeys "{DOWN}"

Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub Form_Load()
AdoArt.CommandType = adCmdText
AdoArt.ConnectionString = cCadConex
AdoArt.RecordSource = "SELECT consec, descripc, precio, str(contenid,10,3) as contenid, medida FROM despensa,tfproduc WHERE clave = consec AND despensa = 0 ORDER BY descripc"
AdoArt.Refresh
Tabla = Caja
End Sub

Private Sub txtNumDesp_GotFocus()
  txtNumDesp.SelStart = 0
  txtNumDesp.SelLength = 5
End Sub

Private Sub txtNumDesp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub txtNumDesp_LostFocus()
If Not IsNumeric(txtNumDesp.Text) Then Exit Sub

Dim RST As ADODB.Recordset
Set RST = New ADODB.Recordset
RST.Open "SELECT consec, descripc, MEDIDA, str(contenid,10,3) as contenid, precio,cantidad FROM despensa, tfproduc WHERE clave = consec AND despensa = 1", cn, adOpenDynamic, adLockOptimistic, adCmdText
cn.Execute "DELETE FROM " & Tabla
While Not RST.EOF
    cn.Execute "INSERT INTO " & Tabla & "(clave,descripc,contenid,medida,cantidad,precio,importe) VALUES ('" & RST!CONSEC & "','" & RST!descripc & "'," & RST!CONTENID & ",'" & RST!medida & "'," & txtNumDesp.Text * RST!cantidad & "," & RST!PRECIO & "," & RST!PRECIO * (Val(txtNumDesp.Text) * RST!cantidad) & ")"
    RST.MoveNext
Wend

AdoDesp.CommandType = adCmdText
AdoDesp.ConnectionString = cCadConex
AdoDesp.RecordSource = "SELECT * FROM " & Tabla & " ORDER BY descripc"
AdoDesp.Refresh
Me.txtNumprod.Text = AdoDesp.Recordset.RecordCount
lblDispo.Visible = True
nDesp = 263.6 * txtNumDesp.Text
txtNumDesp.Enabled = False
cmdOpcion(0).Enabled = True
cmdOpcion(1).Enabled = True
cmdOpcion(2).Enabled = True
End Sub

Private Sub ticdesp()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
For Each x In Printers
   If x.DeviceName Like "*TICKET*" Then
      lImp = True
      Set Printer = x
      Exit For
   End If
Next x
If lImp = False Then
   If MsgBox("NO ES POSIBLE IMPRIMIR TICKET'S PARA SURTIR" & Chr(13) & Chr(13) & "PORQUE NO EXISTEN IMPRESORAS INSTALADAS" & Chr(13) & "PARA ESTE TIPO DE TICKET'S, DESEAS ENVIARLO A LA IMPRESORA PREDETERMINADA", vbCritical + vbYesNo) = vbNo Then Exit Sub
End If
rs.Open "SELECT * FROM " & Tabla & " ORDER BY descripc", cn, adOpenDynamic, adLockOptimistic, adCmdText
Printer.Print " "
Printer.Print "     VIVERES Y LICORES S.A DE C.V.    "
Printer.Print "CARBONERA 1016 COL. TRINIDAD DE H."
Printer.Print "   OAXACA, OAX. " & date & "  " & Format(Time, "HH:MM:SS")
Printer.Print "PRODUCTOS A INCLUIR EN LA DESPENSA"
Printer.Print "-------------------------------------"
nImporte = 0
While Not rs.EOF
  If rs!cantidad > 0 Then
        cCad = CStr(rs!cantidad) & " PZA " & Trim(rs!descripc)
        'En caso de que sea muy grande la descripcion se imprime en dos lineas
        If Len(Trim(cCad)) > 37 Then
           Printer.Print Mid(cCad, 1, 37)
           Printer.Print Mid(cCad, 38, 24);
        Else
           Printer.Print cCad
        End If
        Printer.CurrentX = 172
        Printer.Print "  " & CStr(rs!CONTENID) & " " & rs!medida
        nImporte = nImporte + rs!importe
        nProd = nProd + 1
  End If
  rs.MoveNext
Wend
Printer.Print "-------------------------------------"
Printer.Print "TOTAL DE PRODUCTOS: " & nProd
Printer.Print "TOTAL A PAGAR: " & Format(Val(Format(Me.lblDispo.Caption, "###,###,##0.00")) * -1, "$ ###,###,#00.00")
For N = 0 To 10
   Printer.Print " "
Next
Printer.EndDoc
End Sub
