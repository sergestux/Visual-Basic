VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmHistCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de creditos"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "frmHistCred.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   6360
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "                                                                             Para salir presione la tecla   [Esc]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   11359
            MinWidth        =   11359
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoHisCre 
      Height          =   330
      Left            =   7200
      Top             =   -120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "AdoHisCre"
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
   Begin VB.Frame fraHisCre 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmdAboPar 
         Caption         =   "&Abono Parcial"
         Height          =   300
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Abona un pago al total de la factura"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CmdPagTot 
         Caption         =   "&Pago total"
         Height          =   300
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cambia la situación de la factura a COBRADA"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5400
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5400
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Saldo     :"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Ejercido              :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Plazo     :"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbletiquetas 
         BackColor       =   &H80000004&
         Caption         =   "Limite de credito :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   5175
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   9015
      Begin VB.CommandButton cmdregresar 
         Caption         =   "&Regresar"
         Height          =   300
         Left            =   7560
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCveClie 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   5535
      End
      Begin MSDataGridLib.DataGrid DbgrdHisVen 
         Bindings        =   "frmHistCred.frx":0442
         Height          =   3975
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1.5
         RowHeight       =   15
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "factura"
            Caption         =   "Factura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "noventa"
            Caption         =   "Venta"
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
            DataField       =   "facfecha"
            Caption         =   "Fecha Fact."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "rfc"
            Caption         =   "Rfc"
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
            DataField       =   "total"
            Caption         =   "Total"
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
         BeginProperty Column05 
            DataField       =   "depto1"
            Caption         =   "Depto1"
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
         BeginProperty Column06 
            DataField       =   "depto2"
            Caption         =   "Depto2"
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
         BeginProperty Column07 
            DataField       =   "depto3"
            Caption         =   "Depto3"
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
         BeginProperty Column08 
            DataField       =   "depto4"
            Caption         =   "Depto4"
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
         BeginProperty Column09 
            DataField       =   "depto5"
            Caption         =   "Depto5"
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
         BeginProperty Column10 
            DataField       =   "depto6"
            Caption         =   "Depto6"
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
         BeginProperty Column11 
            DataField       =   "depto7"
            Caption         =   "Depto7"
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
               Alignment       =   1
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmHistCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub adofacturas_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
   cmdAboPar.Enabled = Not AdoFacturas.Recordset!cobrado
End Sub

Private Sub cmbCliente_GotFocus()
 RESP = SendMessageLong(cmbCliente.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbCliente_LostFocus()
 RESP = SendMessageLong(cmbCliente.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbCliente_Validate(Cancel As Boolean)
Dim rscli As ADODB.Recordset
If Trim(cmbCliente.Text) <> "" Then
   Set rscli = New ADODB.Recordset
   ccVeCli = cmbCliente.Text
   rscli.Open "SELECT * FROM Catcliente WHERE cnombre = '" & ccVeCli & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rscli.BOF And rscli.EOF Then
      cmbCliente.SetFocus
      Exit Sub
   End If
   txtCveClie.Text = rscli!cclave
   txtCveClie.SetFocus
Else
   txtCveClie.Text = ""
End If
End Sub

Private Sub cmdAboPar_Click()
  abonoparcial
End Sub

Private Sub CmdPagTot_Click()
Dim RESP, SERIE, Factura
RESP = MsgBox("Esta seguro de Cambiar la situacion de la Factura a cobrada? ", vbYesNo + vbQuestion, "COBRAR")
If RESP = vbYes Then
   nPos = InStr(1, AdoHisCre.Recordset!Factura, "-")
   SERIE = Mid(AdoHisCre.Recordset!Factura, 1, nPos - 1)
   Factura = Trim(Mid(AdoHisCre.Recordset!Factura, nPos + 1))
   cn.Execute "UPDATE facventa SET cobrado = 1 , faccobro = '" & Date & "' WHERE numfactura = '" & Trim(Factura) & "'  AND  serie = '" & SERIE & "'"
   AdoHisCre.Refresh
   MsgBox "Proceso realizado correctamente", vbInformation
End If
End Sub

Private Sub cmdRegresar_Click()
Unload Me
End Sub

Private Sub abonoparcial()
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open "SELECT * FROM CATCLIENTE WHERE ccredito = 1 OR ctipo = 1", cn, adOpenDynamic, adLockOptimistic, adCmdText
  cmbCliente.Clear
  While Not rs.EOF
      cmbCliente.AddItem rs!cNombre
      rs.MoveNext
  Wend
End Sub


Private Sub txtCveClie_LostFocus()
Dim rsvencre As ADODB.Recordset
   Set rs = New ADODB.Recordset
   ccVeCli = String(4 - (Len(Trim(Me.txtCveClie.Text))), "0") + Trim(Me.txtCveClie.Text)
   rs.Open "SELECT * FROM Catcliente WHERE cClave = '" & ccVeCli & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
   If rs.BOF And rs.EOF Then
      cmbCliente.SetFocus
      Exit Sub
   End If
   cmbCliente.Text = rs!cNombre
   txtCveClie.Text = rs!cclave
   
   Set rsvencre = New ADODB.Recordset
   rsvencre.Open "SELECT SUM(Total) AS VentCred FROM facventa WHERE faccliente = '" & ccVeCli & "' AND cobrado = 0 AND cancelado = 0", cn, adOpenKeyset, adLockOptimistic, adCmdText
   VenCre = IIf(IsNull(rsvencre!VentCred), 0, rsvencre!VentCred)
   nDispo = rs!CLIMITECREDITO - VenCre   'Disponible para comprar
   lbletiquetas(4).Caption = Format(rs!CLIMITECREDITO, "###,###,##0.00")
   lbletiquetas(7).Caption = IIf(IsNull(rs!ctiempocredito), 0, rs!ctiempocredito) & " Dias"
   lbletiquetas(5).Caption = Format(VenCre, "###,###,##0.00")
   frmHistCred.lbletiquetas(6).Caption = Format(rs!CLIMITECREDITO - VenCre, "###,###,##0.00")
   
   AdoHisCre.ConnectionString = cCadConex
   AdoHisCre.CommandType = adCmdText
   AdoHisCre.RecordSource = "SELECT RTRIM(serie)+ '-' + LTRIM(numfactura) as factura, noventa,facfecha,total,rfc,depto1,depto2,depto3,depto4,depto5,depto6,depto7 FROM facventa WHERE faccliente = '" & ccVeCli & "' AND cobrado = 0 AND cancelado = 0 ORDER BY facfecha DESC"
   AdoHisCre.Refresh
End Sub
