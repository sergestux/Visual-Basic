VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpedAba 
   Caption         =   "Menu de pedidos por tienda para abastecimiento"
   ClientHeight    =   8595
   ClientLeft      =   255
   ClientTop       =   435
   ClientWidth     =   11880
   Icon            =   "frmpedAba.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoPedidos 
      Height          =   330
      Left            =   1320
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoPedidos"
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
   Begin MSAdodcLib.Adodc AdoDbf 
      Height          =   330
      Left            =   3960
      Top             =   0
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
      CommandType     =   1
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
      Connect         =   "DSN=PITICODBF"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PITICODBF"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM PEDIND"
      Caption         =   "AdoDbf Pedido"
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
   Begin VB.PictureBox PicInf 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11820
      TabIndex        =   14
      Top             =   7650
      Width           =   11880
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   6
         Left            =   5160
         Picture         =   "frmpedAba.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Reporte detallado del pedido para abastecimiento"
         Top             =   90
         Width           =   700
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   5
         Left            =   6000
         Picture         =   "frmpedAba.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Reporte global de pedidos en el rango especificado"
         Top             =   90
         Width           =   700
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   4
         Left            =   4320
         Picture         =   "frmpedAba.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Buscar clave del pedido en el rango seleccionado"
         Top             =   90
         Width           =   700
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   3
         Left            =   3480
         Picture         =   "frmpedAba.frx":1018
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ultimo"
         Top             =   90
         Width           =   700
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   2
         Left            =   2640
         Picture         =   "frmpedAba.frx":118A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Siguiente"
         Top             =   90
         Width           =   700
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   1
         Left            =   1800
         Picture         =   "frmpedAba.frx":12FC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Anterior"
         Top             =   90
         Width           =   700
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   375
         Index           =   0
         Left            =   960
         Picture         =   "frmpedAba.frx":146E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Primero"
         Top             =   90
         Width           =   700
      End
      Begin VB.PictureBox cRpt 
         Height          =   480
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   28
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de pedidos: 999,999"
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
         Left            =   7920
         TabIndex        =   15
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame fradescripcion 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   11655
      Begin MSComDlg.CommonDialog Cmdlg 
         Left            =   2040
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   0
         Left            =   8400
         TabIndex        =   26
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   61079555
         CurrentDate     =   37257
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   1
         Left            =   10080
         TabIndex        =   27
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   61079555
         CurrentDate     =   37257
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Fecha Ini. Imp."
         Height          =   200
         Index           =   0
         Left            =   8160
         TabIndex        =   24
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Fecha Fin.Imp"
         Height          =   195
         Index           =   1
         Left            =   10080
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblSucur 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   285
         Width           =   4695
      End
      Begin VB.Label lblSuc 
         Caption         =   "Sucursal:"
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
         Left            =   600
         TabIndex        =   12
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame cmdDespla 
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Cancelar"
         Height          =   400
         Index           =   6
         Left            =   3720
         Picture         =   "frmpedAba.frx":15E0
         TabIndex        =   5
         ToolTipText     =   "Cancela el pedido seleccionado"
         Top             =   220
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Importa Ped."
         Height          =   400
         Index           =   4
         Left            =   2520
         Picture         =   "frmpedAba.frx":18EA
         TabIndex        =   4
         ToolTipText     =   "Importar pedidos enviados por tiendas para abastecimiento"
         Top             =   220
         Width           =   1095
      End
      Begin VB.ComboBox cmbOrden 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Regresar"
         Height          =   315
         Index           =   5
         Left            =   10320
         Picture         =   "frmpedAba.frx":1BF4
         TabIndex        =   7
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   3
         Left            =   5760
         Picture         =   "frmpedAba.frx":1EFE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Recibir pedido y afectar inventario"
         Top             =   120
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   2
         Left            =   6360
         Picture         =   "frmpedAba.frx":2208
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Confirmar pedido y prepararlo para recibirlo"
         Top             =   120
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Modificar"
         Height          =   400
         Index           =   1
         Left            =   1320
         Picture         =   "frmpedAba.frx":2512
         TabIndex        =   2
         ToolTipText     =   "Modificar pedido capturado"
         Top             =   220
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Nuevo"
         Height          =   400
         Index           =   0
         Left            =   120
         Picture         =   "frmpedAba.frx":281C
         TabIndex        =   1
         ToolTipText     =   "Capturar un nuevo pedido"
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label lblFiltro 
         Alignment       =   2  'Center
         Caption         =   "Filtro"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdPed 
      Bindings        =   "frmpedAba.frx":2B26
      Height          =   5745
      Left            =   150
      TabIndex        =   11
      Top             =   1800
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10134
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1.5
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "p_pedido"
         Caption         =   "FOLIO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "p_proveedor"
         Caption         =   "CVE. PROV."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "p_sucursal"
         Caption         =   "CVE. SUC."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "p_fecped"
         Caption         =   "FECHA  ELABORACION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "p_fecconfirma"
         Caption         =   "        FEC. DE CONF."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "p_fecent"
         Caption         =   "FEC. DE ENVIO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "p_traslado"
         Caption         =   "TRASLADO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "p_cancelado"
         Caption         =   "CANCELADO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "SI"
            FalseValue      =   "NO"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   25
      Top             =   8265
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "                                                  Click en el encabezado ordena los datos en base a la columna   "
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmpedAba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cCond As String     'Condicion del filtro del grid
Private ccondrpt As String  'Condicion del filtro del rpt
Private rstSucProv As ADODB.Recordset
Private ntext As Integer
Private cFecha As String
Private cFecharpt As String

Private Sub AdoPedidos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
   lblSucur.Caption = AdoPedidos.Recordset!tidescrip
End Sub

Private Sub cmborden_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub cmborden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub

Private Sub cmborden_LostFocus()
Select Case cmbOrden.ListIndex
Case 0 'Todos
    cCond = "P_proveedor = 'ABA'"
    ccondrpt = "{PEDIDOS.P_Proveedor} = 'ABA'"
Case 1 'Pendientes por Confirmar
    cCond = "P_situacion = 0"
    ccondrpt = "{PEDIDOS.P_situacion} = 0"
Case 2 'Pendientes por recibir
    cCond = "P_situacion = 1  AND P_recibido = 0"
    ccondrpt = "{PEDIDOS.P_situacion} = 1  AND {PEDIDOS.P_recibido} = 0"
Case 3 'Recibidos
    cCond = "P_recibido = 1"
    ccondrpt = "{PEDIDOS.P_recibido} = 1"
End Select

cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
cFecha = " AND (month(p_fecped) >= " & Month(dtpFecha(0).Value) & " and (day(p_fecped) > = " & Day(dtpFecha(0).Value) & cOper & " (day(p_fecped)<= " & Day(dtpFecha(1).Value) & " and month(p_fecped)<= " & Month(dtpFecha(1).Value) & ")) and year(p_fecped)>= " & Year(dtpFecha(0).Value) & " and year(p_fecped)<= " & Year(dtpFecha(1).Value) & ")"

'AdoPedidos.RecordSource = "SELECT * FROM [Pedidos] WHERE " & cCond & cFecha
AdoPedidos.RecordSource = "SELECT * FROM Pedidos, CatTienda WHERE p_sucursal = ticlave AND " & cCond & cFecha
AdoPedidos.Refresh
For N = 0 To 6
    Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
Next
cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
cmdOpcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
'cmdOpcion(4).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
lblInfo.Caption = "Numero de pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
End Sub

Private Sub cmdMoverse_Click(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0  'Primer registro
    AdoPedidos.Recordset.MoveFirst
Case 1  'Anterior
    AdoPedidos.Recordset.MovePrevious
    If AdoPedidos.Recordset.BOF Then AdoPedidos.Recordset.MoveFirst
Case 2  'Siguiente
    AdoPedidos.Recordset.MoveNext
    If AdoPedidos.Recordset.EOF Then AdoPedidos.Recordset.MoveLast
Case 3  'Ultimo
    AdoPedidos.Recordset.MoveLast
Case 4  'Buscar Clave de pedido
    cCve = InputBox("Introduzca la clave del pedido a buscar", "Introducir clave")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    Antes = dbgrdPed.Bookmark
    AdoPedidos.Recordset.MoveFirst
    AdoPedidos.Recordset.Find "p_pedido = '" & Trim(cCve) & "'"
    If AdoPedidos.Recordset.EOF Then
        MsgBox "LA CLAVE " & cCve & " NO SE ENCUENTRA EN LOS PEDIDOS " + IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text), vbExclamation
        dbgrdPed.Bookmark = Antes
    End If
Case 5
    cMensaje = stb1.SimpleText
    stb1.SimpleText = Space(120) + "Espere un momento generando reporte................"
    stb1.Refresh
    cFecharpt = " AND {PEDIDOS.P_fecped} >= Date(" & CStr(Year(dtpFecha(0).Value)) & "," & CStr(Month(dtpFecha(0).Value)) & "," & CStr(Day(dtpFecha(0).Value)) & ") AND {PEDIDOS.P_fecped} <= Date(" & CStr(Year(dtpFecha(1).Value)) & "," & CStr(Month(dtpFecha(1).Value)) & "," & CStr(Day(dtpFecha(1).Value)) & ")"
    crpt.ReportFileName = App.Path & "\Pedidos.rpt"
    crpt.WindowTitle = "Reporte de pedidos"
    crpt.Formulas(0) = "FORMSELEC = " & ccondrpt & cFecharpt
    crpt.Formulas(1) = "PEDIDO = 'LISTADO DE PEDIDOS " & IIf(cmbOrden.Text = "TODOS", "", cmbOrden.Text) & " DEL " & Trim(dtpFecha(0).Value) & " AL " & Trim(dtpFecha(1).Value) & " '"
    crpt.Connect = cCadConex
    crpt.Action = 1
    stb1.SimpleText = cMensaje
    stb1.Refresh
Case 6
    mensaje = stb1.SimpleText
    stb1.SimpleText = Space(65) & "Espere un momento generando reporte"
    stb1.Refresh
    'cRpt.SelectionFormula = "{PEDIDOS.p_pedido} = '" & AdoPedidos.Recordset!p_pedido & "'"
    crpt.Connect = cCadConex
    crpt.ReportFileName = App.Path & "\PeCapCon.rpt"
    crpt.WindowTitle = "Pedido para abastecimiento numero " & AdoPedidos.Recordset!p_Pedido
    crpt.Formulas(0) = "FORMSELEC = '" & AdoPedidos.Recordset!p_Pedido & "'"
    crpt.Formulas(1) = "PEDIDO = 'PEDIDO PARA ABASTECIMIENTO CON FOLIO [ " & AdoPedidos.Recordset!p_Pedido & " ]'"
    crpt.Formulas(2) = "ENCA = 'PEDIDO PARA SUCURSAL [ " & lblSucur.Caption & " ]'"
    'Es mucho mas rapido que el selectionformula
    'cRpt.SQLQuery = "SELECT PEDIDOS.p_pedido, PEDIDOS.p_proveedor, PEDIDOS.p_fecped, PEDIDOS.p_sucursal, PEDIDOS.p_fecentreal, PEDIDOS.p_fecconfirma, " & _
    '                        " DETALLEFACTURA.df_prod, DETALLEFACTURA.df_cantidad, DETALLEFACTURA.df_cantsol, DETALLEFACTURA.df_cantsolp, " & _
    '                        " CATPROV.NOMPROVE, CATTIENDA.tidescrip, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES " & Chr(13) & _
    '                "FROM pitico.dbo.PEDIDOS PEDIDOS, " & _
    '                     "pitico.dbo.DETALLEFACTURA DETALLEFACTURA," & _
    '                     "pitico.dbo.CATPROV CATPROV," & _
    '                     "pitico.dbo.CATTIENDA CATTIENDA," & _
    '                     "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
    '                "WHERE PEDIDOS.p_pedido = DETALLEFACTURA.df_pedido AND " & _
    '                     "PEDIDOS.p_proveedor = CATPROV.PROVE AND " & _
    '                     "PEDIDOS.p_sucursal = CATTIENDA.ticlave AND " & _
    '                     "DETALLEFACTURA.df_prod = TFPRODUC.CONSEC  AND " & _
    '                     "AND " & _
    '                     "PEDIDOS.p_pedido = '" & AdoPedidos.Recordset!p_pedido & "' " & Chr(13) & _
    '                "ORDER BY " & _
    '                     "TFPRODUC.DESCRIPC ASC, TFPRODUC.CONTENID ASC"
    'MsgBox cRpt.Formulas(0)
    crpt.Action = 1
    stb1.SimpleText = mensaje
    stb1.Refresh
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
'On Error GoTo Error:
Select Case Index
  Case 0  'Nuevo pedido
       cModo = "CAPTURARPEDIDO"
       nOp = 1
       frmCaptPed.Caption = "Nuevo pedido para abastecimiento"
       frmCaptPed.Show
  Case 1  'Modificar pedido
       cModo = "CAPTURARPEDIDO"
       nOp = 0
       frmCaptPed.Caption = "Modificar pedido para abastecimiento"
       frmCaptPed.Show
       SendKeys AdoPedidos.Recordset!p_Pedido
       SendKeys "{TAB}"
  Case 2  'Confirmar pedido
       cModo = "CONFIRMARPEDIDO"
       nOp = 0
       frmCaptPed.Caption = "Confirmar pedido"
       frmCaptPed.Show
       SendKeys frmpedidos.dbgrdPed.Columns(0).Text
       SendKeys "{TAB}"
  Case 3  'Recibir pedido
       cModo = "RECIBIRPEDIDO"
       nOp = 0
       frmCaptPed.Caption = "Modificar pedido"
       frmCaptPed.Show
       SendKeys frmpedidos.dbgrdPed.Columns(0).Text
       SendKeys "{TAB}"
  Case 4 'Importar pedidos de tiendas para abastecimiento
         'Importar pedidos de tabla Dbf de tiendas a Sql Server tabla PEDIDOS y DETALLEFACTURAS
       ImpPedAba '"IND"
  Case 5  'Salir del modulo de pedidos
       Unload Me
  Case 6
       If MsgBox("REALMENTE DESEAS CANCELAR EL PEDIDO CON FOLIO: " & AdoPedidos.Recordset!p_Pedido, vbInformation + vbYesNo) = vbYes Then
          AdoPedidos.Recordset!P_CANCELADO = 1
          AdoPedidos.Recordset.Update
       End If
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdopcion_GotFocus(Index As Integer)
If Index = 0 Then Unload frmAreaRecibo
End Sub

Private Sub dbgrdPed_DblClick()
  If Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF) Then cmdOpcion_Click 1
End Sub

Private Sub DbGrdped_HeadClick(ByVal ColIndex As Integer)
  stb1.SimpleText = Space(65) + "Espere un momento ordenando Pedidos por " & dbgrdPed.Columns(ColIndex).Caption
  AdoPedidos.RecordSource = "SELECT * FROM [Pedidos] WHERE " & cCond & cFecha & "ORDER BY " & dbgrdPed.Columns(ColIndex).DataField
  AdoPedidos.Refresh
  stb1.SimpleText = Space(85) + "Pedidos ordenandos por " & dbgrdPed.Columns(ColIndex).Caption
End Sub

Private Sub dbgrdPed_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 If dbgrdPed.SelBookmarks.Count > 0 Then dbgrdPed.SelBookmarks.Remove 0
 dbgrdPed.SelBookmarks.Add dbgrdPed.RowBookmark(dbgrdPed.Row)
End Sub

Private Sub dtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Activate()
  Unload frmAreaRecibo
  Forma = 0    'Se activa bandera de pedidos sugeridos
  If dtpFecha(0).Value = "01/01/02" Then dtpFecha(0).Value = date
  If dtpFecha(1).Value = "01/01/02" Then dtpFecha(1).Value = date
  
  cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
  cFecha = " AND (month(p_fecped) >= " & Month(dtpFecha(0).Value) & " and (day(p_fecped) > = " & Day(dtpFecha(0).Value) & cOper & " (day(p_fecped)<= " & Day(dtpFecha(1).Value) & " and month(p_fecped)<= " & Month(dtpFecha(1).Value) & ")) and year(p_fecped)>= " & Year(dtpFecha(0).Value) & " and year(p_fecped)<= " & Year(dtpFecha(1).Value) & ")"  'Cargo todos los pedidos
  
  AdoPedidos.ConnectionString = cCadConex
  AdoPedidos.CommandType = adCmdText
  AdoPedidos.RecordSource = "SELECT * FROM Pedidos, CatTienda WHERE p_sucursal = ticlave AND " & cCond & cFecha
  AdoPedidos.Refresh
  For N = 0 To 6
     Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  Next
  cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
  cmdOpcion(4).Enabled = (tipotienda = 2 Or tipotienda = 4)
  lblInfo.Caption = "Numero de pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
  'Obtengo datos para el nombre del proveedor y la sucursal
  Set rstSucProv = New ADODB.Recordset
  rstSucProv.CursorType = adOpenDynamic
  rstSucProv.Source = "SELECT PEDIDOS.p_pedido, PEDIDOS.p_proveedor,PEDIDOS.p_sucursal, CATTIENDA.ticlave, CATTIENDA.tidescrip FROM PEDIDOS, CATTIENDA WHERE PEDIDOS.p_sucursal = CATTIENDA.ticlave"
  rstSucProv.ActiveConnection = cCadConex
  rstSucProv.Open
  
  cmbOrden.AddItem "TODOS"
  cmbOrden.ListIndex = 0
  
  cCond = "p_proveedor = 'ABA'"               ' Filtro por default todos los pedidos
  ccondrpt = "{PEDIDOS.p_proveedor} = 'ABA' AND {DETALLEFACTURA.df_sugerido} = 0"  ' Filtro por default del RPT
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmAreaRecibo.Show
End Sub

Private Sub ImpPedAba() '(CveTie As String)
On Error GoTo Error:
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim Archivo  As String
Dim SUC As String
   MenAnt = stb1.SimpleText
   Cmdlg.FileName = ""
   Cmdlg.CancelError = True
   Cmdlg.DialogTitle = "Abrir archivo enviado por tiendas para abastecimiento"
   Cmdlg.Filter = "Archivos de Pedidos (Ped???.txt) | Ped???.txt"
   Cmdlg.ShowOpen
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If

   Archivo = Mid(Cmdlg.FileName, Len(Cmdlg.FileName) - 9)
   SUC = Pedsuc(Archivo)
   If IsNull(SUC) Or Trim(SUC) = "" Then
      MsgBox "EL NOMBRE DEL ARCHIVO: " & cArch & "NO ESTA REGISTRADO EN EL SISTEMA" & Chr(13) _
             & "Y NO SE LE ASIGNARA SUCURSAL AL PEDIDO POR LO TANTO SE CANCELARA LA IMPORTACION" & Chr(13) & "FAVOR DE AVISAR AL ADMINISTRADOR DEL SISTEMA", vbCritical
      Exit Sub
   End If
   
  'Obtengo el folio mayor de la tienda
   Set rsttemp = New ADODB.Recordset
   rsttemp.Open "SELECT MAX (CAST(SUBSTRING(P_PEDIDO,4,7) AS INT)) As FolMay FROM [PEDIDOS] WHERE SUBSTRING(p_pedido,1,3) = '" & Mid(Trim(Archivo), 4, 3) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If IsNull(rsttemp!FolMay) Then
      FolPed = UCase(Mid(Trim(Archivo), 4, 3)) & "1"
   Else
      FolPed = UCase(Mid(Trim(Archivo), 4, 3)) & Trim(Str(rsttemp!FolMay + 1))
   End If

   cn.Execute "INSERT INTO Pedidos(p_pedido,p_proveedor,p_fecped,p_sucursal) SELECT Ped ='" & FolPed & "', Prove = 'ABA', Fechasol = '" & date + Time & "',Suc = '" & SUC & "'"
   'Agrego el detalle de Factura
   rsttemp.Close
   Open cRutArc For Input As #1
   While Not EOF(1)
       Line Input #1, CAD
       pos1 = InStr(CAD, "|")
       clave = Mid(CAD, 1, pos1 - 1)
       CAD = Mid(CAD, pos1 + 1, Len(CAD))
       pos1 = InStr(CAD, "|")
       cajas = Mid(CAD, 1, pos1 - 1)
       CAD = Mid(CAD, pos1 + 1, Len(CAD))
       piezas = Val(CAD)

      stb1.SimpleText = Space(55) & "Importando producto: " & clave
      stb1.Refresh
      rsttemp.Open "SELECT * FROM TFPRODUC WHERE CONSEC = '" & clave & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
      If rsttemp.RecordCount > 0 Then
         rsttemp.Close
         rsttemp.Open "SELECT Costo As Pre FROM DESCPROD WHERE PRODUCTO = '" & CONSEC & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
         nprecio = IIf(rsttemp.RecordCount > 1, rsttemp!pre, 0)
         cn.Execute "INSERT INTO DetalleFactura(df_prod,df_pedido,df_cantsol,df_cantsolp,df_costo,df_sugerido) Select Prod = '" & clave & "', Ped = '" & FolPed & "', CantSolC = " & cajas & ", CantSolPza = " & piezas & ",Precio = " & CStr(nprecio) & ", Abaste = 0 "
      Else
         MsgBox "EL PRODUCTO CON LA CLAVE " & clave & " NO EXISTE EN EL CATALOGO "
      End If
      rsttemp.Close
   Wend
   MsgBox "SE GENERO EL PEDIDO CON FOLIO: " & FolPed & Chr(13) & " Y FECHA " & date + Time, vbInformation
   AdoPedidos.Refresh
   stb1.SimpleText = MenAnt
   stb1.Refresh
  Exit Sub
Error:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
   End If

End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub dtpFecha_LostFocus(Index As Integer)
On Error Resume Next
 cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
 cFecha = " AND (month(p_fecped) >= " & Month(dtpFecha(0).Value) & " and (day(p_fecped) > = " & Day(dtpFecha(0).Value) & cOper & " (day(p_fecped)<= " & Day(dtpFecha(1).Value) & " and month(p_fecped)<= " & Month(dtpFecha(1).Value) & ")) and year(p_fecped)>= " & Year(dtpFecha(0).Value) & " and year(p_fecped)<= " & Year(dtpFecha(1).Value) & ")"

 'AdoPedidos.RecordSource = "SELECT * FROM [Pedidos] WHERE " & cCond & cFecha & " ORDER BY p_fecped"
 AdoPedidos.RecordSource = "SELECT * FROM Pedidos, CatTienda WHERE p_sucursal = ticlave AND " & cCond & cFecha
 AdoPedidos.Refresh
 lblInfo.Caption = "Numero de Pedidos:" + Str(AdoPedidos.Recordset.RecordCount)
 For N = 0 To 6   'Si esta vacio el recordset desactivo las opciones
   Cmdmoverse(N).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 Next
 cmdOpcion(1).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 cmdOpcion(2).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)
 cmdOpcion(3).Enabled = Not (AdoPedidos.Recordset.BOF And AdoPedidos.Recordset.EOF)

End Sub
