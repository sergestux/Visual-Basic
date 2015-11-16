VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCodBarra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar etiquetas de codigo de barras"
   ClientHeight    =   7950
   ClientLeft      =   1410
   ClientTop       =   1890
   ClientWidth     =   11400
   Icon            =   "FrmcodBarra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraImp 
      Caption         =   "Imprimir Etiquetas"
      Height          =   3975
      Left            =   1680
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox TxtTextoeti 
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   25
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txtCampos 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   7
         Left            =   6120
         TabIndex        =   26
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox cmbPuerto 
         Height          =   315
         ItemData        =   "FrmcodBarra.frx":0442
         Left            =   840
         List            =   "FrmcodBarra.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Regresar"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   30
         Top             =   3360
         Width           =   1695
      End
      Begin VB.ComboBox cmbCodigos 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   720
         Width           =   7575
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   1680
         TabIndex        =   27
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Numero de etiquetas a Imprimir:"
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   33
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Puerto de Impresora"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   32
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "ESPECIFICACIONES DEL PRODUCTO"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   7575
      End
      Begin VB.Label lblNumEti 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   7215
      End
   End
   Begin VB.Frame FraGen 
      Caption         =   "Generar etiquetas"
      Height          =   3735
      Left            =   840
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox txtCampos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/mm/yy h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   4
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbUsuario 
         Height          =   315
         Left            =   2760
         TabIndex        =   20
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Regresar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmbProd 
         Height          =   315
         Left            =   2760
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.ComboBox cmbproved 
         Height          =   315
         Left            =   2760
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   6735
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Fecha elab."
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Numero de cajas"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Producto"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbletiquetas 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport cRpt 
      Left            =   120
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowLeft      =   0
      WindowTop       =   0
      WindowState     =   2
   End
   Begin MSCommLib.MSComm Com 
      Left            =   240
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7575
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "                                                                                            Para salir presione la tecla  [ESC]"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraMenu 
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Command1"
         Height          =   495
         Index           =   1
         Left            =   7320
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Ver Eti. Gen."
         Height          =   500
         Index           =   5
         Left            =   4440
         TabIndex        =   12
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Reimprimir"
         Height          =   495
         Index           =   4
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Imprimir"
         Height          =   500
         Index           =   2
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1250
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Salir"
         Height          =   500
         Index           =   3
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   1250
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Generar"
         Height          =   500
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1250
      End
   End
   Begin MSAdodcLib.Adodc adocajas 
      Height          =   330
      Left            =   7560
      Top             =   2640
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
      Caption         =   "Adodc1"
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
   Begin VB.Frame FraReimp 
      Caption         =   "Reimprimir eiquetas"
      Height          =   1455
      Left            =   3000
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbPuertoR 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtReimp 
         Height          =   375
         Left            =   3480
         TabIndex        =   35
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Puerto de Impresora"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo de la etiqueta a imprimir"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.Frame FraVerEti 
      Caption         =   "Ver etiquetas impresas"
      Height          =   6375
      Left            =   840
      TabIndex        =   39
      Top             =   1200
      Visible         =   0   'False
      Width           =   9975
      Begin MSACAL.Calendar Cal1 
         Height          =   1935
         Left            =   3600
         TabIndex        =   47
         Top             =   4320
         Visible         =   0   'False
         Width           =   2655
         _Version        =   524288
         _ExtentX        =   4683
         _ExtentY        =   3413
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2000
         Month           =   6
         Day             =   14
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   2
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtfecha 
         Height          =   375
         Left            =   1560
         TabIndex        =   46
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   495
         Left            =   8280
         Picture         =   "FrmcodBarra.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5640
         Width           =   1455
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "Reporte"
         Height          =   495
         Left            =   6720
         Picture         =   "FrmcodBarra.frx":05CE
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   5640
         Width           =   1335
      End
      Begin VB.ListBox lstEti 
         Height          =   1035
         Left            =   480
         TabIndex        =   42
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dbgrdVerEti 
         Bindings        =   "FrmcodBarra.frx":0B00
         Height          =   4335
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Codigo"
            Caption         =   "CODIGO"
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
            DataField       =   "DESCRIPC"
            Caption         =   "DESCRIPCION"
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
            DataField       =   "MEDIDA"
            Caption         =   "ESPECIF."
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
            DataField       =   "NUMETI"
            Caption         =   "ETI. GEN."
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
            DataField       =   "ETIIMP"
            Caption         =   "ETI. IMP."
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
               Button          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3869.858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1590.236
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1019.906
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoVerEti 
         Height          =   375
         Left            =   960
         Top             =   1920
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
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
         Caption         =   "Adodc1"
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
         Caption         =   "Fecha"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   45
         Top             =   5880
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCodBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsttemp As ADODB.Recordset
Private rstProd As ADODB.Recordset
Private cProv As String
Private cProd As String

Private Sub Cal1_DblClick()
   txtFecha.Text = Cal1.Value
   dbgrdVerEti.Caption = "ETIQUETAS GENERADAS EL DIA: " & UCase(Format(Date, "LONG DATE"))
   AdoVerEti.ConnectionString = cCadConex
   AdoVerEti.CommandType = adCmdText
   AdoVerEti.RecordSource = "SELECT COUNT(*) AS NUMETI , COUNT(FECHAIMPRESION) AS etiimp, SUBSTRING(CODIGO,1,10) AS CODIGO, DESCRIPC, STR(PAQUETES) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' + TFPRODUC.MEDIDA AS MEDIDA FROM CODIGOS,TFPRODUC WHERE TFPRODUC.CONSEC = SUBSTRING(CODIGO,1,10) AND YEAR(FECHACREACION)= " & Year(txtFecha.Text) & " And Month(FECHACREACION) = " & Month(txtFecha.Text) & " AND DAY(FECHACREACION)= " & Day(txtFecha.Text) & " GROUP BY SUBSTRING(CODIGO,1,10),DESCRIPC,PAQUETES,CONTENID,MEDIDA ORDER BY descripc"
   AdoVerEti.Refresh
End Sub

Private Sub cmbCodigos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub cmbCodigos_LostFocus()
If Trim(cmbCodigos.Text) <> "" Then
    rsttemp.Close
    lblNumEti.Visible = True
    rsttemp.Open "SELECT COUNT(CODIGO) AS NumCod FROM CODIGOS WHERE CODIGOS.FECHAIMPRESION IS NULL AND SUBSTRING(CODIGO,1,10) = '" & Mid(cmbCodigos.Text, 1, 10) & "'"
    lblNumEti.Caption = "NUMERO DE ETIQUETAS PENDIENTES DE IMPRIMIRSE: " & CStr(rsttemp!NumCod)
    txtCampos(7).Text = rsttemp!NumCod
    rsttemp.Close
    rsttemp.Open "SELECT DESCRIPC FROM TFPRODUC WHERE CONSEC = '" & Mid(cmbCodigos.Text, 1, 10) & "'"
    TxtTextoeti.Text = Trim(Mid(rsttemp!Descripc, 1, 25))
End If
End Sub

Private Sub cmbProd_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
     Case 27
         FraGen.Visible = False
         ActDesOpc True
     Case 13
         KeyAscii = 0
         SendKeys "{Tab}"
End Select
End Sub

Private Sub cmbProd_Validate(Cancel As Boolean)
On Error GoTo Error:
If cmbProd.Text = "" Or IsNull(cmbProd.Text) Then
   MsgBox "Debe seleccionar un producto de la lista desplegable", vbExclamation
   cmbProd.SetFocus
   Cancel = True
Else
   rstProd.MoveFirst
   rstProd.Find "CONSEC = '" & Trim(Mid(cmbProd.Text, Len(cmbProd.Text) - 10)) & "'"
   If rstProd.EOF = True Then
      MsgBox "Debe seleccionar un producto de la lista desplegable", vbExclamation
      cmbProd.SetFocus
      Cancel = True
   Else
   txtCampos(1).Text = rstProd!CONSEC
   txtCampos(1).SetFocus
   End If
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmbproved_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
     Case 27
          FraGen.Visible = False
          ActDesOpc True
     Case 13
         KeyAscii = 0
         SendKeys "{Tab}"
End Select

End Sub

Private Sub cmbproved_Validate(Cancel As Boolean)

If cmbproved.Text = "" Or IsNull(cmbproved.Text) Then
   MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
   cmbproved.SetFocus
   Cancel = True
Else
   rsttemp.MoveFirst
   rsttemp.Find "NomProve = '" & cmbproved.Text & "'"
   If rsttemp.EOF = True Then
      MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
      cmbproved.SetFocus
      Cancel = True
   Else
   txtCampos(0).Text = rsttemp!Prove
   txtCampos(0).Enabled = True
   txtCampos(0).SetFocus
   End If
End If
End Sub

Private Sub cmdCerrar_Click(Index As Integer)
  Select Case Index
  Case 0
       FraGen.Visible = False
  Case 1
       Me.FraImp.Visible = False
  End Select
  ActDesOpc True
End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
  ActDesOpc True
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo Error:
Dim rstCod As ADODB.Recordset
If Not IsNumeric(txtCampos(2).Text) Then
   MsgBox "LA CANTIDAD DE CAJAS DEBE SER NUMERICA", vbExclamation
   Exit Sub
End If
If txtCampos(2).Text < 1 Then
   MsgBox "LA CANTIDAD DE CAJAS DEBE SER MAYOR A CERO", vbExclamation
   Exit Sub
End If

rsttemp.Close
rsttemp.Open "SELECT MAX (CAST(SUBSTRING(Codigo,11,15) AS INT)) As FolMay FROM CODIGOS WHERE SUBSTRING(Codigo,1,10) = '" & cProd & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
UltNum = IIf(IsNull(rsttemp!FolMay), 0, rsttemp!FolMay)
Set rstCod = New ADODB.Recordset
rstCod.Open "CODIGOS", cn, adOpenKeyset, adLockOptimistic, adCmdTable
For n = 1 To Val(txtCampos(2).Text)
    rstCod.AddNew
    rstCod!Codigo = cProd + CStr(UltNum + n)
    rstCod!FECHACREACION = txtCampos(4).Text
    rstCod!Fechaimpresion = Null
    rstCod!Entradainv = 0
    rstCod!Salidainv = 0
    rstCod!Usuario = txtCampos(3).Text
    rstCod.Update
Next
MsgBox "LAS ETIQUETAS SE GENERARON CORRECTAMENTE", vbInformation
FraGen.Visible = False
ActDesOpc True
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdImprimir_Click()
Dim rstInv As ADODB.Recordset
On Error GoTo Error:
rsttemp.Close
rsttemp.Open "SELECT COUNT(CODIGO) AS NumCod FROM CODIGOS WHERE CODIGOS.FECHAIMPRESION IS NULL AND SUBSTRING(CODIGO,1,10) = '" & Mid(cmbCodigos.Text, 1, 10) & "'"

If Trim(cmbCodigos.Text) = "" Then
   MsgBox "DEBE DE ESPECIFICAR ETIQUETA DEL PRODUCTO A IMPRIMIR", vbExclamation
   Exit Sub
ElseIf Not IsNumeric(txtCampos(7).Text) Then
   MsgBox "EL NUMERO DE ETIQUETAS A IMPRIMIR DEBE SER NUMERICO", vbExclamation
   Exit Sub
ElseIf Val(txtCampos(7).Text) <= 0 Then
   MsgBox "EL NUMERO DE ETIQUETAS DEBE SER MAYOR A CERO", vbExclamation
   Exit Sub
ElseIf Val(txtCampos(7).Text) > rsttemp!NumCod Then
   MsgBox "EL NUMERO DE ETIQUETAS A IMPRIMIRSE DEBE SER MENOR" & Chr(13) & "O IGUAL A LAS PENDIENTES DE IMPRIMIRSE", vbExclamation
   Exit Sub
End If
cResp = MsgBox("DESEAS IMPRIMIR " & txtCampos(7).Text & " ETIQUETAS E ICREMENTAR EL INVENTARIO?" & Chr(13) & _
          "REALMENTE DESEAS IMPRIMIR LAS ETIQUETAS", vbQuestion + vbYesNo)
If cResp = vbYes Then
    rsttemp.Close
    rsttemp.Open "SELECT * FROM CODIGOS,TFPRODUC WHERE CODIGOS.FECHAIMPRESION IS NULL AND SUBSTRING(CODIGO,1,10) = TFPRODUC.CONSEC AND SUBSTRING(CODIGO,1,10) = '" & Mid(cmbCodigos.Text, 1, 10) & "' ORDER BY CONVERT(INT,SUBSTRING(CODIGO,10,5))"
           
    Com.CommPort = cmbPuerto.ListIndex + 1
    Com.Settings = "9600,N,8,1"
    Com.PortOpen = True
    ' Se establecen propiedades de etiquetas
    Com.Output = "{F,1,A,R,M,0762,0508," & Chr(34) & "ONLINE" & Chr(34) & "|"
    Com.Output = "T,001,21,V,0502,0153,0,2,1,1,B,L,0,3|"
    Com.Output = "T,002,15,V,0546,0051,0,1,1,1,B,L,0,3|"
    Com.Output = "B,003,15,V,0226,0331,8,8,0127,8,L,1|"
    Com.Output = "T,004,30,V,0721,0407,0,1,1,1,B,L,0,3|"
    Com.Output = "T,005,22,V,0657,0359,0,1,1,1,B,L,0,3|"
    Com.Output = "T,006,08,V,0200,0153,0,2,1,1,B,L,0,3|"
    Com.Output = "}"

    Set rstInv = New ADODB.Recordset
    rstInv.Open "SELECT * FROM Inventario WHERE Inprod = '" & Trim(Mid(rsttemp!Codigo, 1, 10)) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If (rstInv.BOF And rstInv.EOF) Then
       MsgBox "NO EXISTE EN EL INVENTARIO EL ARTICULO " & Chr(13) & _
              rsttemp!CONSEC & "  " & rsttemp!Descripc & Chr(13) & CStr(rsttemp!Paquetes) & " X " & CStr(rsttemp!Contenid) & " " & rsttemp!Medida & Chr(13) & _
              "A CONTINUACION SE DARA DE ALTA EN EL INVENTARIO", vbInformation
       rstInv.AddNew
       rstInv!Inprod = rsttemp!CONSEC
       rstInv!Insucursal = Mid(cSucursal, 1, 3)
       rstInv!InObserva = " "
       rstInv!InFecCaduProx = "1/1/1900"
       rstInv!instock = 0
       rstInv!inInicial = 0
       rstInv!Incant = 0
       rstInv.Update
    End If

    For n = 1 To Val(txtCampos(7).Text) 'Numero de etiquetas a imprimirse
        'se envia a impresion una etiqueta
        If Com.InBufferCount = 0 Then
            cadena = CStr(rsttemp!Paquetes) & " EN " & CStr(rsttemp!Contenid) & " " & rsttemp!Medida
            Com.Output = "{B,1,N,0001|"
            Com.Output = "001," & Chr(34) & rsttemp!Codigo & Chr(34) & "|"                'Texto que aparece abajo del codigo de barras
            Com.Output = "002," & Chr(34) & Mid(cadena, 1, 15) & Chr(34) & "|" 'Ultima linea de la etiqueta
            Com.Output = "003," & Chr(34) & rsttemp!Codigo & Chr(34) & "|"                'Texto que aparece abajo del codigo de barras
            Com.Output = "004," & Chr(34) & Mid(Trim(TxtTextoeti.Text), 1, 17) & Chr(34) & "|"  'Primera linea
            Com.Output = "005," & Chr(34) & Mid(Trim(TxtTextoeti.Text), 18, 8) & Chr(34) & "|" 'Primera linea
            Com.Output = "006," & Chr(34) & Date & Chr(34) & "|"  'Primera linea
            Com.Output = "}"
        
            'Grabo fecha, hora de impresion, entrada al inventario y fecha de entrada.
            cn.Execute "UPDATE CODIGOS SET FechaImpresion = '" & Date + Time & "', EntradaInv = '1', FechaEntrada = '" & Date + Time & "', Producto = '" & rsttemp!CONSEC & "' WHERE codigo = '" & rsttemp!Codigo & "'"
            'Actualizo el inventario al imprimirse la etiqueta
            rstInv!Incant = rstInv!Incant + (1 * rsttemp!Paquetes)
            rstInv.Update
            rsttemp.MoveNext
        Else
            If Com.InBufferCount > 1 Then
               Com.InBufferCount = 0
            End If
            n = n - 1
        End If
   Next
   Com.OutBufferCount = 0
   Com.InBufferCount = 0
   Com.PortOpen = False
End If
If Com.PortOpen = True Then
 Com.OutBufferCount = 0
 Com.InBufferCount = 0
 Com.PortOpen = False
End If

Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub generadbf()
dbf.Recordset.AddNew
dbf.Recordset!Codigo = rsttemp!Codigo
dbf.Recordset!Paquetes = CStr(rsttemp!Paquetes) & " X " & CStr(rsttemp!Contenid) & " " & rsttemp!Medida
dbf.Recordset!Descripc = rsttemp!Descripc
dbf.Recordset!Fecha = Date
dbf.Recordset.Update
End Sub
Private Sub cmdopcion_Click(Index As Integer)

Select Case Index
Case 0  'Generar etiquetas

    ActDesOpc False 'Procedimiento que desactiva la opcion de menus
    txtCampos(3).Text = Mid(cCveDesUsu, 1, 3)
    cmbUsuario.Text = Trim(Mid(cCveDesUsu, 3))
    txtCampos(3).Enabled = False
    cmbUsuario.Enabled = False
    Me.FraGen.Visible = True
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT * FROM Catprov ORDER BY NomProve", cn, adOpenKeyset, adLockOptimistic, adCmdText
    cMensaje = stb1.SimpleText
    stb1.SimpleText = Space(75) & "Espere un momento cargando catalogo de proveedores"
    stb1.Refresh
    cmbproved.Clear
    Do While Not rsttemp.EOF
       If Not IsNull(rsttemp!nomprove) Then
          cmbproved.AddItem rsttemp!nomprove
       End If
       rsttemp.MoveNext
    Loop
    stb1.SimpleText = cMensaje
    stb1.Refresh
    cmbproved.Enabled = True
    txtCampos(0).Enabled = True
    txtCampos(0).SetFocus
Case 2  'Imprimir etiquetas
    ActDesOpc False 'Procedimiento que desactiva la opcion de menus
    FraImp.Visible = True
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT DISTINCT SUBSTRING(CODIGO,1,10) AS ClaProd, TFPRODUC.descripc, LTRIM(STR(TFPRODUC.paquetes,10,3)) + ' X ' +  lTrim(Str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA " & _
                 "FROM CODIGOS,TFPRODUC WHERE TFPRODUC.Consec = SUBSTRING(CODIGO,1,10) AND CODIGOS.FECHAIMPRESION IS NULL", cn, adOpenKeyset, adLockOptimistic, adCmdText
    cmbCodigos.Clear
    Do While Not rsttemp.EOF
       cmbCodigos.AddItem rsttemp!ClaProd & Space(2) & rsttemp!Descripc & Space(2) & rsttemp!Medida
       rsttemp.MoveNext
    Loop
    cmbPuerto.ListIndex = 0
    lblNumEti.Visible = False
Case 3
    Unload Me
Case 4
    Me.FraReimp.Visible = True
    ActDesOpc False
    Me.cmbPuertoR.ListIndex = 0
    MsgBox "AL REIMPRIMIR SOLAMENTE SE REPONE LA ETIQUETA" & Chr(13) & "SE ACTUALIZA LA FECHA DE IMPRESION" & Chr(13) & "PERO NO SE AFECTA EL INVENTARIO", vbExclamation
    txtReimp.SetFocus
Case 5
   ActDesOpc False 'Procedimiento que desactiva la opcion de menus
   FraVerEti.Visible = True
   If txtFecha.Text = "" Then txtFecha.Text = Date
   dbgrdVerEti.Caption = "ETIQUETAS GENERADAS EL DIA: " & UCase(Format(Date, "LONG DATE"))
   AdoVerEti.ConnectionString = cCadConex
   AdoVerEti.CommandType = adCmdText
   AdoVerEti.RecordSource = "SELECT COUNT(*) AS NUMETI , COUNT(FECHAIMPRESION) AS etiimp, SUBSTRING(CODIGO,1,10) AS CODIGO, DESCRIPC, STR(PAQUETES) + ' X ' + LTRIM(STR(CONTENID,10,3)) + ' ' + TFPRODUC.MEDIDA AS MEDIDA FROM CODIGOS,TFPRODUC WHERE TFPRODUC.CONSEC = SUBSTRING(CODIGO,1,10) AND DAY(FECHACREACION)= " & CStr(Day(Date)) & " GROUP BY SUBSTRING(CODIGO,1,10),DESCRIPC,PAQUETES,CONTENID,MEDIDA ORDER BY descripc"
   AdoVerEti.Refresh

End Select
End Sub

Private Sub cmdOpcion_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then

End If
End Sub


Private Sub cmdReimp_Click()
Me.FraReimp.Visible = True
ActDesOpc False
Me.cmbPuertoR.ListIndex = 0
txtReimp.SetFocus
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   FraVerEti.Visible = False
   ActDesOpc True
End If
End Sub


Private Sub cmdRegresar_Click()
   FraVerEti.Visible = False
   ActDesOpc True
End Sub

Private Sub cmdReporte_Click()
    cMensaje = stb1.SimpleText
    stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
    stb1.Refresh
    cFecharpt = " {CODIGOS.fechacreacion} = Date(" & CStr(Year(txtFecha.Text)) & "," & CStr(Month(txtFecha.Text)) & "," & CStr(Day(txtFecha.Text)) & ") "
    cRpt.ReportFileName = App.Path & "\Entradas.rpt"
    cRpt.WindowTitle = "Reporte de entradas"
    cRpt.Formulas(0) = "FORMSELEC =" & cFecharpt
    cRpt.Formulas(1) = "ENCABEZADO = 'ENTRADAS DIARIAS DE CAJAS DEL " & Trim(txtFecha.Text) & "'"
    cRpt.Connect = cCadConex
    cRpt.Action = 1
    stb1.SimpleText = cMensaje
    stb1.Refresh

End Sub

Private Sub dbgrdVerEti_ButtonClick(ByVal ColIndex As Integer)
Dim rsttemp As ADODB.Recordset
On Error GoTo Error:
     'Abajo (3):
    lstEti.Left = dbgrdVerEti.Left + dbgrdVerEti.Columns(0).Left
    lstEti.Top = dbgrdVerEti.Top + dbgrdVerEti.RowTop(dbgrdVerEti.Row) + dbgrdVerEti.RowHeight
          '.Width = dbgrdDetPed.Columns(ColIndex).Width + 15
          '.ListIndex = 0
    lstEti.Visible = True
    lstEti.ZOrder 0
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT * FROM CODIGOS WHERE YEAR(FECHACREACION) = '" & Year(txtFecha.Text) & "' AND MONTH(FECHACREACION) =' " & Month(txtFecha.Text) & "' AND DAY(FECHACREACION) ='" & Day(txtFecha.Text) & "' AND SUBSTRING(CODIGO,1,10) = '" & AdoVerEti.Recordset!Codigo & "' ORDER BY CONVERT(INT,SUBSTRING(CODIGO,10,5))", cn, adOpenDynamic, adLockOptimistic, adCmdText
    lstEti.Clear
    While Not rsttemp.EOF
       lstEti.AddItem rsttemp!Codigo
       rsttemp.MoveNext
    Wend
    lstEti.SetFocus
   Exit Sub
Error:
   MsgBox Err.Description

End Sub

Private Sub dbgrdVerEti_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   FraVerEti.Visible = False
   ActDesOpc True
End If
End Sub

Private Sub dbgrdVerEti_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
lstEti.Visible = False
Cal1.Visible = False
End Sub

Private Sub Form_Activate()
Unload frmpedBod

End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmpedBod.Show
End Sub


Private Sub txtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0  'Clave del Proveedor
     txtCampos(Index).Text = Trim(UCase(txtCampos(Index).Text))
     txtCampos(Index).Refresh
     If txtCampos(Index).Text = "" Or IsNull(txtCampos(Index).Text) Then
        cmbproved.SetFocus
        Exit Sub
     Else
        rsttemp.MoveFirst
        rsttemp.Find "Prove= '" & Trim(txtCampos(Index).Text) & "'"
        If rsttemp.EOF = True Then
           cmbproved.SetFocus
           Exit Sub
        End If
    End If
    cmbProd.Clear
    cmbproved.Text = rsttemp!nomprove
    cProv = rsttemp!Prove
    Set rstProd = New ADODB.Recordset
    rstProd.Open "SELECT * FROM  tfproduc WHERE Claprove = '" & txtCampos(Index).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not rstProd.EOF
       If Not IsNull(rstProd!CONSEC) Then
          cmbProd.AddItem rstProd!Descripc & " " & CStr(rstProd!Paquetes) & " X " & CStr(rstProd!Contenid) & " " & rstProd!Medida & " " & rstProd!CONSEC
       End If
       rstProd.MoveNext
    Loop
    txtCampos(0).Enabled = False
    cmbproved.Enabled = False
    
    LBLETIQUETAS(1).Visible = True
    txtCampos(1).Visible = True
    cmbProd.Visible = True
    txtCampos(1).Text = ""
    txtCampos(1).SetFocus
    cmbProd.SetFocus
Case 1  'Clave del producto
     txtCampos(Index).Text = Trim(UCase(txtCampos(Index).Text))
     txtCampos(Index).Refresh
     If txtCampos(Index).Text = "" Or IsNull(txtCampos(Index).Text) Then
        cmbProd.SetFocus
        Exit Sub
     Else
        rstProd.MoveFirst
        rstProd.Find "CONSEC= '" & Trim(txtCampos(Index).Text) & "'"
        If rstProd.EOF = True Then
           'cmbProd.SetFocus
           Exit Sub
        End If
    End If
    cmbProd.Text = rstProd!Descripc & " " & CStr(rstProd!Paquetes) & " X " & CStr(rstProd!Contenid) & " " & rstProd!Medida & " " & rstProd!CONSEC
    cProd = rstProd!CONSEC
    
    CmdGrabar.Visible = True
    cmdCerrar(0).Visible = True
    LBLETIQUETAS(4).Visible = True
    txtCampos(4).Visible = True
    txtCampos(4).Text = Date + Time
    txtCampos(4).Enabled = False
    LBLETIQUETAS(2).Visible = True
    txtCampos(2).Visible = True
    txtCampos(2).SetFocus
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub ActDesOpc(lvalor As Boolean)
  For n = 0 To 5
     cmdOpcion(n).Enabled = lvalor
  Next
End Sub


Private Sub txtCancela_Change()

End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
   cmdConCance_Click
End If
End Sub
Private Sub etiq()
    Com.CommPort = cmbPuerto.ListIndex + 1
            'Com.Settings = "00,N,8,1,P"
   
           Com.Settings = "9600,N,8,1"
           Com.PortOpen = True
           ' Se establecen propiedades de etiquetas
           Com.Output = "{F,1,A,R,M,0762,0508," & Chr(34) & "ONLINE" & Chr(34) & "|"
           Com.Output = "T,001,21,V,0502,0153,0,2,1,1,B,L,0,3|"
           Com.Output = "T,002,15,V,0546,0051,0,1,1,1,B,L,0,3|"
           Com.Output = "B,003,15,V,0226,0331,8,8,0127,8,L,1|"
           Com.Output = "T,004,30,V,0721,0407,0,1,1,1,B,L,0,3|"
           Com.Output = "T,005,22,V,0657,0359,0,1,1,1,B,L,0,3|"
           Com.Output = "T,006,08,V,0200,0153,0,2,1,1,B,L,0,3|"
           Com.Output = "}"
           Com.InBufferCount = 0
       
   rsttemp.MoveFirst
   For n = 1 To Val(txtCampos(7).Text) 'Numero de etiquetas a imprimirse
        'se envia a impresion una etiqueta
        If Com.InBufferCount = 0 Then
            'Com.InBufferCount
            Com.Output = "{B,1,N,0001|"
            Com.Output = "001," & Chr(34) & rsttemp!Codigo & Chr(34) & "|"                'Texto que aparece abajo del codigo de barras
            Com.Output = "002," & Chr(34) & CStr(rsttemp!Paquetes) & " EN " & CStr(rsttemp!Contenid) & " " & rsttemp!Medida & Chr(34) & "|"   'Ultima linea de la etiqueta
            Com.Output = "003," & Chr(34) & rsttemp!Codigo & Chr(34) & "|"                'Texto que aparece abajo del codigo de barras
            Com.Output = "004," & Chr(34) & Mid(rsttemp!Descripc, 1, 30) & Chr(34) & "|"  'Primera linea
            Com.Output = "005," & Chr(34) & Mid(rsttemp!Descripc, 31) & Chr(34) & "|"  'Primera linea
            Com.Output = "006," & Chr(34) & Date & Chr(34) & "|"  'Primera linea
            Com.Output = "}"
        
            'Grabo fecha y hora de impresion de etiqueta
            cn.Execute "UPDATE CODIGOS SET FechaImpresion = '" & Date + Time & "' WHERE codigo = '" & rsttemp!Codigo & "'"
            'Por levantamiento de inventario al Imprimir se le da entrada al articulo
            'cn.Execute "UPDATE CODIGOS SET EntradaInv = '1',FechaEntrada = '" & Date + Time & "' WHERE codigo = '" & rsttemp!codigo & "'"
            'nCon = nCon + 1
            rsttemp.MoveNext
        Else
            'MsgBox "IN" & CStr(Com.InBufferCount)
            If Com.InBufferCount > 1 Then
               Com.InBufferCount = 0
            End If
            'MsgBox "OUT" & CStr(Com.OutBufferCount)
            'Com.OutBufferCount = 0
            'Com.OutBufferCount
            n = n - 1
        End If
   Next
   Com.OutBufferCount = 0
   Com.InBufferCount = 0
   Com.PortOpen = False
If Com.PortOpen = True Then
 Com.OutBufferCount = 0
 Com.InBufferCount = 0
 Com.PortOpen = False
End If
End Sub

Private Sub txtFecha_GotFocus()
  Cal1.Value = txtFecha.Text
  Cal1.Visible = True
End Sub

Private Sub txtReimp_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then  'salir
       FraReimp.Visible = False
       ActDesOpc True
  ElseIf KeyAscii = 13 Then   'Reimprimir
       If Len(txtReimp.Text) <> 15 Then
           Me.txtReimp.SetFocus
           Exit Sub
       End If
       'Busco la etiqueta
       Set rsttemp = New ADODB.Recordset
       rsttemp.Open "SELECT * FROM CODIGOS WHERE Codigo = '" & txtReimp.Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
       If rsttemp.BOF And rsttemp.EOF Then
            MsgBox "LA ETIQUETA CON LA CLAVE:" & txtReimp.Text & " NO EXISTE", vbExclamation
            txtReimp.Text = ""
            txtReimp.SetFocus
            Exit Sub
       Else
           cmbPuertoR.ListIndex = 0
           rsttemp.Close
           rsttemp.Open "SELECT * FROM CODIGOS,TFPRODUC WHERE SUBSTRING(CODIGO,1,10) = TFPRODUC.CONSEC AND CODIGOS.Codigo = '" & txtReimp.Text & "'"
           cResp = MsgBox("DESEAS REIMPRIMIR ETIQUETA DEL SIGUIENTE PRODUCTO" & Chr(13) & _
                    rsttemp!CONSEC & " " & rsttemp!Descripc & Chr(13) & CStr(rsttemp!Paquetes) & " EN " & CStr(rsttemp!Contenid) & " " & rsttemp!Medida, vbInformation + vbYesNo)
           If cResp = vbYes Then
                Com.CommPort = cmbPuertoR.ListIndex + 1
                Com.Settings = "9600,N,8,1"
                Com.PortOpen = True
                
                If Com.InBufferCount > 1 Then
                   Com.InBufferCount = 0
                End If
                ' Se establecen propiedades de etiquetas
                Com.Output = "{F,1,A,R,M,0762,0508," & Chr(34) & "ONLINE" & Chr(34) & "|"
                Com.Output = "T,001,21,V,0502,0153,0,2,1,1,B,L,0,3|"
                Com.Output = "T,002,15,V,0546,0051,0,1,1,1,B,L,0,3|"
                Com.Output = "B,003,15,V,0226,0331,8,8,0127,8,L,1|"
                Com.Output = "T,004,30,V,0721,0407,0,1,1,1,B,L,0,3|"
                Com.Output = "T,005,22,V,0657,0359,0,1,1,1,B,L,0,3|"
                Com.Output = "T,006,08,V,0200,0153,0,2,1,1,B,L,0,3|"
                Com.Output = "}"

                'se envia a impresion una etiqueta
                cadena = CStr(rsttemp!Paquetes) & " EN " & CStr(rsttemp!Contenid) & " " & rsttemp!Medida
                Com.Output = "{B,1,N,0001|"
                Com.Output = "001," & Chr(34) & rsttemp!Codigo & Chr(34) & "|"                'Texto que aparece abajo del codigo de barras
                Com.Output = "002," & Chr(34) & Mid(cadena, 1, 15) & Chr(34) & "|" 'Ultima linea de la etiqueta
                Com.Output = "003," & Chr(34) & rsttemp!Codigo & Chr(34) & "|"                'Texto que aparece abajo del codigo de barras
                Com.Output = "004," & Chr(34) & Mid(Trim(rsttemp!Descripc), 1, 17) & Chr(34) & "|"  'Primera linea
                Com.Output = "005," & Chr(34) & Mid(Trim(rsttemp!Descripc), 18, 8) & Chr(34) & "|" 'Primera linea
                Com.Output = "006," & Chr(34) & Date & Chr(34) & "|"  'Primera linea
                Com.Output = "}"
        
                'Grabo fecha y hora de impresion de etiqueta
                cn.Execute "UPDATE CODIGOS SET FechaImpresion = '" & Date + Time & "' WHERE codigo = '" & rsttemp!Codigo & "'"
                Com.PortOpen = False
           End If
       End If
End If
End Sub

Private Sub TxtTextoeti_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub
