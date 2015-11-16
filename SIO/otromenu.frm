VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdesp 
   AutoRedraw      =   -1  'True
   Caption         =   "Ventas de Productos"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "otromenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR1 
      Left            =   4440
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Fraopciones 
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   11655
      Begin VB.ComboBox cmbsucursal 
         Height          =   315
         ItemData        =   "otromenu.frx":030A
         Left            =   120
         List            =   "otromenu.frx":0332
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmdconsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   10320
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker FechaFinal 
         Height          =   375
         Left            =   8280
         TabIndex        =   19
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421376
         CalendarTitleForeColor=   16777215
         Format          =   59965441
         CurrentDate     =   37064
      End
      Begin MSComCtl2.DTPicker FechaIni 
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421376
         CalendarTitleForeColor=   -2147483634
         Format          =   59965441
         CurrentDate     =   37064
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Fecha Inicial"
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Fecha Final"
         Height          =   255
         Left            =   8280
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fra1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton cmdopcion 
         Caption         =   "&Verificar Dias"
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   3
         Left            =   6240
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Exporta a &Excel"
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Importa &ventas"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Importa &Mes"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Data ventasdbf 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adoventas 
      Height          =   330
      Left            =   120
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   0
      CommandTimeout  =   5
      CursorType      =   3
      LockType        =   1
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adoventas"
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
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   11775
      Begin VB.TextBox txtbusca 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   5760
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Frame fraDias 
         BackColor       =   &H8000000C&
         Height          =   5295
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   5415
         Begin MSDataGridLib.DataGrid dbgrdDias 
            Bindings        =   "otromenu.frx":03A7
            Height          =   4335
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   7646
            _Version        =   393216
            AllowUpdate     =   -1  'True
            ForeColor       =   8388608
            HeadLines       =   1.3
            RowHeight       =   15
            RowDividerStyle =   0
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "FECHA"
               Caption         =   "                Día"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dddddd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "ventas"
               Caption         =   "     Ventas"
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
            BeginProperty Column02 
               DataField       =   "fecha"
               Caption         =   "Fecha"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   3
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               RecordSelectors =   0   'False
               BeginProperty Column00 
                  DividerStyle    =   0
                  ColumnWidth     =   2564.788
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  DividerStyle    =   0
                  ColumnWidth     =   1110.047
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   989.858
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdDias 
            Caption         =   "&Regresar"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   4920
            Width           =   5055
         End
      End
      Begin ComctlLib.ProgressBar probar1 
         Height          =   255
         Left            =   5640
         TabIndex        =   11
         Top             =   5760
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog Cmdlg 
         Left            =   1200
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dbgrdVentas 
         Bindings        =   "otromenu.frx":03BF
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   9340
         _Version        =   393216
         HeadLines       =   1.8
         RowHeight       =   15
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "proveedor"
            Caption         =   "        Proveedor"
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
            DataField       =   "producto"
            Caption         =   "                                      Producto"
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
            DataField       =   "paquetes"
            Caption         =   "Paq."
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
            Caption         =   "Contenid."
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
         BeginProperty Column05 
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
         BeginProperty Column06 
            DataField       =   "costo"
            Caption         =   "Cost. Prom."
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
            DataField       =   "venta"
            Caption         =   "Pre. Prom."
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
            DataField       =   "linea"
            Caption         =   "Linea"
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
         BeginProperty Column09 
            DataField       =   "Familia"
            Caption         =   "Familia"
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
         BeginProperty Column10 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2340.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5504.882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1110.047
            EndProperty
         EndProperty
      End
      Begin VB.Label LblEti 
         Alignment       =   2  'Center
         Caption         =   "Búsqueda por descripción del producto"
         Height          =   255
         Left            =   5640
         TabIndex        =   26
         Top             =   5550
         Width           =   6015
      End
      Begin VB.Label lblImporte 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0.00"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label lblEtiquetas 
         Caption         =   "IMPORTE"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label LBLREG 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "VAR. PRODUCTOS:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   5760
         Width           =   1575
      End
   End
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   8265
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "                                                                                          Para salir presione la tecla [ Esc ]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmdesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdconsultar_Click()
LBLREG.Visible = True
Label4.Visible = True
CADENA = sucursalfile()
If CADENA = "" And cmbsucursal.Text <> "_TODAS" Then
   MsgBox "Necesita seleccionar la sucursal a Importar...", vbCritical
   Exit Sub
End If
If Trim(cmbsucursal.Text) = "_TODAS" Then
   If tipotienda = 1 Then
      cadstring = "SELECT max(nomprove) proveedor, max(descripc) AS producto , max(contenid) contenid , max(medida) medida, max(paquetes) paquetes , " & _
            " MAX(V.CONSEC) CONSEC, SUM(V.CANTIDAD) CANTIDAD , AVG(V.PREVENTA) VENTA, AVG(V.PRECOSTO) COSTO, SUM(IMPORTE) IMPORTE, MAX(T.LINEA) as LINEA, MAX(L.SFFAMILIA) AS FAMILIA, MAX(BARRASPZA) AS BARRAS" & _
            " FROM ventasacum1 v, tfproduc t, catprov p, lineas l WHERE l.sfclave = t.linea AND p.prove = t.claprove and v.consec = t.consec" & _
            " AND V.FECHA >= '" & FechaIni.Value & "' AND V.FECHA <= '" & FechaFinal.Value & "'" & _
            " GROUP BY V.CONSEC order by producto "
   Else
      cadstring = "SELECT max(nomprove) proveedor, max(descripc) AS producto , max(contenid) contenid , max(medida) medida, max(paquetes) paquetes , " & _
            " MAX(V.CONSEC) CONSEC, SUM(V.CANTIDAD) CANTIDAD , AVG(V.PREVENTA) VENTA, AVG(V.PRECOSTO) COSTO, SUM(IMPORTE) IMPORTE, MAX(T.LINEA) as LINEA, FAMILIA = "", MAX(BARRASPZA) AS BARRAS" & _
            " FROM VENGRAL v, tfproduc t, catprov p WHERE p.prove = t.claprove and v.consec = t.consec" & _
            " AND V.FECHA >= '" & FechaIni.Value & "' AND V.FECHA <= '" & FechaFinal.Value & "'" & _
            " GROUP BY V.CONSEC order by producto "
   End If
Else
   sucu = Mid(CADENA, 4, 3)
   If tipotienda = 1 Then
      cadstring = "select max(nomprove) proveedor, max(descripc) AS producto , max(contenid) contenid , max(medida) medida, max(paquetes) paquetes , " & _
            " MAX(V.CONSEC) CONSEC, SUM(V.CANTIDAD) CANTIDAD , AVG(V.PREVENTA) VENTA, AVG(V.PRECOSTO) COSTO, SUM(IMPORTE) IMPORTE, MAX(T.LINEA) as LINEA, MAX(L.SFFAMILIA) AS FAMILIA, MAX(BARRASPZA) AS BARRAS" & _
            " FROM ventasacum1 v, tfproduc t, catprov p, lineas l WHERE l.sfclave = t.linea AND p.prove = t.claprove and v.consec = t.consec and sucursal = '" & sucu & "'" & _
            " AND V.FECHA >= '" & FechaIni.Value & "' AND V.FECHA <= '" & FechaFinal.Value & "'" & _
            " GROUP BY V.CONSEC order by producto "
   Else
      cadstring = "select max(nomprove) proveedor, max(descripc) AS producto , max(contenid) contenid , max(medida) medida, max(paquetes) paquetes , " & _
            " MAX(V.CONSEC) CONSEC, SUM(V.CANTIDAD) CANTIDAD , AVG(V.PREVENTA) VENTA, AVG(V.PRECOSTO) COSTO, SUM(IMPORTE) IMPORTE, MAX(T.LINEA) as LINEA, FAMILIA = "", MAX(BARRASPZA) AS BARRAS" & _
            " FROM vengral v, tfproduc t, catprov p WHERE p.prove = t.claprove and v.consec = t.consec and sucursal = '" & sucu & "'" & _
            " AND V.FECHA >= '" & FechaIni.Value & "' AND V.FECHA <= '" & FechaFinal.Value & "'" & _
            " GROUP BY V.CONSEC order by producto "
   End If
End If
MenAnt = Stb1.SimpleText
Stb1.SimpleText = Space(40) & "Espere un momento consultando ventas del período seleccionado....."
Stb1.Refresh

Adoventas.RecordSource = cadstring
Adoventas.Refresh
If Adoventas.Recordset.RecordCount = 0 Then
   MsgBox "NO EXISTEN VENTAS DE " & cmbsucursal.Text & " EN EL PERIODO SELECCIONADO", vbExclamation, "Módulo de compras"
   Exit Sub
End If

LBLREG.Caption = Adoventas.Recordset.RecordCount
LBLREG.Refresh
nImporte = 0
While Not Adoventas.Recordset.EOF
   nImporte = nImporte + IIf(IsNull(Adoventas.Recordset!importe), 0, Adoventas.Recordset!importe)
   lblImporte.Caption = Format(nImporte, "$###,###,##0.00")
   lblImporte.Refresh
   Adoventas.Recordset.MoveNext
Wend
Stb1.SimpleText = MenAnt
Stb1.Refresh
txtbusca.Visible = Not (Adoventas.Recordset.BOF And Adoventas.Recordset.EOF)
End Sub

Private Sub cmdDias_Click()
  Me.fraDias.Visible = False
  dbgrdVentas.Visible = True
End Sub

Private Sub cmdopcion_Click(Index As Integer)
txtbusca.Visible = False
Select Case Index
 Case 0
      Call importaventas
 Case 1
      Call ImpVtaDia
 Case 2
     res = MsgBox("Debe tener seleccionado el periodo y la Sucursal a procesar  , Desea Continuar... ", vbYesNo + vbInformation)
     If res = vbYes Then
        Call expventas
     End If
 Case 3
     Unload Me
 Case 4
     LBLREG.Visible = False
     Label4.Visible = False
     If Me.cmbsucursal.Text = "" Then
        MsgBox "ES NECESARIO ESPECIFICAR LA SUCURSAL DE LA QUE SE VERIFICARAN VENTAS", vbExclamation
        Exit Sub
     End If
     cMens = Me.Stb1.SimpleText
     Stb1.SimpleText = Space(20) & "Espere un momento consultando dias de los que existe registro de ventas"
     dbgrdVentas.Visible = False
     fraDias.Visible = True
     If cmbsucursal.Text = "_TODAS" Then
        If tipotienda = 1 Then
           Adoventas.RecordSource = "SELECT fecha as FECHA, SUM(importe) as VENTAS FROM ventasacum1 WHERE FECHA >= '" & FechaIni.Value & "' AND FECHA <= '" & FechaFinal.Value & "' GROUP BY fecha ORDER BY FECHA"
        Else
           Adoventas.RecordSource = "SELECT fecha as FECHA, SUM(precio) as VENTAS FROM vengral WHERE FECHA >= '" & FechaIni.Value & "' AND FECHA <= '" & FechaFinal.Value & "' GROUP BY fecha ORDER BY FECHA"
        End If
     Else
        CADENA = sucursalfile()
        If tipotienda = 1 Then
           Adoventas.RecordSource = "SELECT fecha as FECHA, SUM(importe) as VENTAS FROM ventasacum1 WHERE sucursal = '" & Mid(CADENA, 4, 3) & "' AND FECHA >= '" & FechaIni.Value & "' AND FECHA <= '" & FechaFinal.Value & "' GROUP BY fecha ORDER BY FECHA"
        Else
           Adoventas.RecordSource = "SELECT fecha as FECHA, SUM(precio) as VENTAS FROM vengral WHERE FECHA >= '" & FechaIni.Value & "' AND FECHA <= '" & FechaFinal.Value & "' GROUP BY fecha ORDER BY FECHA"
        End If
     End If
     Adoventas.Refresh
     monto = 0
     While Not Adoventas.Recordset.EOF
        monto = monto + Adoventas.Recordset!ventas
        Adoventas.Recordset.MoveNext
     Wend
     lblImporte.Caption = Format(monto, "$###,###,##0.00")
     If Not (Adoventas.Recordset.BOF And Adoventas.Recordset.EOF) Then Adoventas.Recordset.MoveFirst
     Stb1.SimpleText = cMens
     Stb1.Refresh
End Select
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub drv1_Change()
  Dir1.Path = drv1.Drive
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then
   frmCalc.Show 1
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Unload frmAreaRecibo
Adoventas.CursorType = adOpenStatic
Adoventas.LockType = adLockOptimistic
Adoventas.CommandType = adCmdText
Adoventas.ConnectionString = strconnect
Adoventas.CommandTimeout = 0
Adoventas.ConnectionTimeout = 0
'Adoventas.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=pitico;Data Source=SERVIDOR_OAXACA"
'Adoventas.RecordSource = "select * from pventas"
'Adoventas.Refresh
FechaIni.Value = Date
Me.FechaFinal = Date
cmdopcion(0).Enabled = (tipotienda = 1)
cmdopcion(1).Enabled = (tipotienda = 1)
End Sub

Public Sub SELECTFILE()
On Error GoTo Error:
Cmdlg.FileName = ""
 Cmdlg.CancelError = True
 Cmdlg.DialogTitle = "Nombre del Archivo de Excel"
 Cmdlg.Filter = "Archivos Excel (*.xls) | *.xls"
 Cmdlg.ShowOpen
 cArch = Cmdlg.FileName
 If cArch = "" Or IsNull(cArch) Then
    MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
    Exit Sub
 End If
 Exit Sub
Error:
  MsgBox "Es necesario Especificar un nombre de archivo", vbCritical
End Sub

Private Sub expventas()
Call SELECTFILE
MsgBox "SE PROCESARAN " & Adoventas.Recordset.RecordCount & " REGISTROS... ", vbInformation
MenAnt = Stb1.SimpleText
Stb1.SimpleText = Space(40) & "Espere un momento generando el archivo en excel....."
Stb1.Refresh


Adoventas.Recordset.MoveFirst
'CONFIGURACION DEL ARCHIVO DE EXCEL
Dim Targetdir As String
Contador_Project_Planning = 1
Contador_Project_Selected = 1
Contador_Project_Files = 1
MousePointer = vbHourglass
Targetdir = cArch
Set Appxl = CreateObject("Excel.Application")  'run it
Appxl.Visible = True
Appxl.Workbooks.Add
If Adoventas.Recordset.BOF And Adoventas.Recordset.EOF Then
       MsgBox "No existen Productos con existencias ... ", vbInformation
Else
     probar1.Visible = True
     probar1.Enabled = True
     probar1.Min = 0
     probar1.Max = Adoventas.Recordset.RecordCount
     contador_Fila = 1
     MousePointer = vbHourglass
     ' Renombro las hojas de Excel y le pongo los encabezados para introducir los datos
     Appxl.Sheets("hoja1").Select
     Appxl.Sheets("hoja1").Name = "VENTAS"
     Appxl.Cells(Contador_Project_Selected, 1).Value = "PROVEEDOR"
     Appxl.Cells(Contador_Project_Selected, 2).Value = "PRODUCTO"
     Appxl.Cells(Contador_Project_Selected, 3).Value = "PAQUETES"
     Appxl.Cells(Contador_Project_Selected, 4).Value = "CONTENIDO"
     Appxl.Cells(Contador_Project_Selected, 5).Value = "MEDIDA"
     Appxl.Cells(Contador_Project_Selected, 6).Value = "PIEZAS"
     Appxl.Cells(Contador_Project_Selected, 7).Value = "PRE. COSTO"
     Appxl.Cells(Contador_Project_Selected, 8).Value = "PRE. VTA"
     Appxl.Cells(Contador_Project_Selected, 9).Value = "IMPORTE VTA."
     Appxl.Cells(Contador_Project_Selected, 10).Value = "LINEA"
     Appxl.Cells(Contador_Project_Selected, 11).Value = "FAMILIA"
     Appxl.Cells(Contador_Project_Selected, 12).Value = "BARRAS"
     Appxl.Cells(Contador_Project_Selected, 13).Value = "CLAVE"
     Contador_Project_Selected = Contador_Project_Selected + 1
     Appxl.Sheets("VENTAS").Select
    While Not Adoventas.Recordset.EOF
        Appxl.Cells(Contador_Project_Selected, 1).Value = Adoventas.Recordset!PROVEEDOR
        Appxl.Cells(Contador_Project_Selected, 2).Value = Adoventas.Recordset!producto
        Appxl.Cells(Contador_Project_Selected, 3).Value = Adoventas.Recordset!PAQUETES
        Appxl.Cells(Contador_Project_Selected, 4).Value = Adoventas.Recordset!Contenid
        Appxl.Cells(Contador_Project_Selected, 5).Value = Adoventas.Recordset!medida
        Appxl.Cells(Contador_Project_Selected, 6).Value = Adoventas.Recordset!cantidad
        Appxl.Cells(Contador_Project_Selected, 7).Value = Adoventas.Recordset!costo
        Appxl.Cells(Contador_Project_Selected, 8).Value = Adoventas.Recordset!venta
        Appxl.Cells(Contador_Project_Selected, 9).Value = Adoventas.Recordset!importe
        Appxl.Cells(Contador_Project_Selected, 10).Value = Adoventas.Recordset!linea
        Appxl.Cells(Contador_Project_Selected, 11).Value = Adoventas.Recordset!familia
        Appxl.Cells(Contador_Project_Selected, 12).Value = Adoventas.Recordset!Barras
        Appxl.Cells(Contador_Project_Selected, 13).Value = Adoventas.Recordset!CONSEC
        Contador_Project_Selected = Contador_Project_Selected + 1
        v = v + 1
        probar1.Value = v
        Me.LBLREG.Caption = v
        LBLREG.Refresh
        Adoventas.Recordset.MoveNext
    Wend
    Adoventas.Recordset.Close
    Call ordenaventas
 End If
 
'SE CIERRA EXCEL
    'appXL.ActiveWorkbook.Save Targetdir
    Appxl.ActiveWorkbook.SaveAs Targetdir
    'Appxl.ActiveWorkbook.Close (False)
    'appxl.Application.Quit
    Set Appxl = Nothing
    MousePointer = vbDefault
    Stb1.SimpleText = MenAnt
    Stb1.Refresh
    MsgBox "PROCESO FINALIZADO...", vbInformation
    probar1.Visible = False
    probar1.Enabled = False
    Exit Sub
End Sub

Private Sub ordenaventas()
' Macro1 Macro
' Macro grabada el 30/04/2001 por Moises Leon
'    Cells.Select
'    With Selection.Font
'        .Name = "Arial"
'        .Size = 8
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ColorIndex = xlAutomatic
'    End With
'    Range("F5").Select
'    Columns("C:C").EntireColumn.AutoFit
'    Columns("D:D").EntireColumn.AutoFit
'    Columns("E:E").EntireColumn.AutoFit
'    Columns("G:G").Select
'    Columns("F:F").EntireColumn.AutoFit
'    Columns("G:G").EntireColumn.AutoFit
'    Columns("H:H").EntireColumn.AutoFit
'    Range("D3").Select
'    Columns("A:A").EntireColumn.AutoFit
'    Columns("B:B").EntireColumn.AutoFit
'    Columns("A:A").ColumnWidth = 38.14
'    Columns("B:B").ColumnWidth = 34.86
'    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Key2:=Range("B2") _
'        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
'        False, Orientation:=xlTopToBottom
'    ActiveWorkbook.Save
'    Selection.AutoFilter
'    ActiveWorkbook.Save
End Sub

Private Function sucursalfile()
Select Case cmbsucursal.Text
Case "PERIFERICO"
     sucursalfile = "VENPER"
Case "REFORMA"
     sucursalfile = "VENREF"
Case "BRENAMIEL"
     sucursalfile = "VENBRE"
Case "CONZATI"
     sucursalfile = "VENCON"
Case "DOLORES"
     sucursalfile = "VENDOL"
Case "JPGARCIA"
     sucursalfile = "VENJPG"
Case "MINA"
     sucursalfile = "VENMIN"
Case "ROSARIO"
     sucursalfile = "VENROS"
Case "SAN MARTIN"
     sucursalfile = "VENSMA"
Case "HIDALGO"
     sucursalfile = "VENHID"
Case "CENTRAL"
     sucursalfile = "VENCEN"
End Select
End Function

Private Sub importaventas()
ventasdbf.Connect = "dbase III;"
MsgBox "El archivo debe estar en p:\ventas", vbInformation
ventasdbf.DatabaseName = "p:\ventas"
ventasdbf.RecordsetType = Table
CADENA = sucursalfile()
If CADENA = "" Then
   MsgBox "Necesita seleccionar la sucursal a Importar...", vbCritical
   Exit Sub
End If
ventasdbf.RecordSource = CADENA

ventasdbf.Refresh
LBLREG.Caption = ventasdbf.Recordset.RecordCount
'MsgBox ventasdbf.Recordset.RecordCount
If ventasdbf.Recordset.RecordCount > 0 Then
    probar1.Enabled = True
    probar1.Visible = True
    probar1.Min = 0
    probar1.Max = ventasdbf.Recordset.RecordCount
    probar1.Visible = True
    v = 0
    ventasdbf.Recordset.MoveFirst
    If ventasdbf.Recordset!fecha = "01/01/01" Then
        MsgBox "Este archivo ya fue importado", vbCritical
        Exit Sub
    End If
    Me.MousePointer = 11
    While Not ventasdbf.Recordset.EOF
         fecha = ventasdbf.Recordset!fecha
         CONSEC = ventasdbf.Recordset!CONSEC
         cantidad = IIf(ventasdbf.Recordset!cantidad > 0, ventasdbf.Recordset!cantidad, 0)
         PRECOSTO = IIf(ventasdbf.Recordset!PRECOSTO > 0, ventasdbf.Recordset!PRECOSTO, 0)
         Preventa = IIf(ventasdbf.Recordset!Preventa > 0, ventasdbf.Recordset!Preventa, 0)
         'PRECIO = IIf(ventasdbf.Recordset!importe > 0, ventasdbf.Recordset!importe, 0)
         PRECIO = IIf(ventasdbf.Recordset!importe > 0, ventasdbf.Recordset!importe, cantidad * Preventa)
         estacion = Trim(ventasdbf.Recordset!Caja)
         sucursal = Mid(CADENA, 4, 3)
         'se contruye la sentencia
         cad = "INSERT INTO ventasacum1(fecha,consec,cantidad,IMPORTE,precosto,preventa,sucursal,estacion)" & _
         " values('" & fecha & "','" & Trim(CONSEC) & "'," & cantidad & "," & PRECIO & "," & PRECOSTO & "," & Preventa & ",'" & sucursal & "','" & estacion & "')"
         cn.Execute cad
         ventasdbf.Recordset.Edit
         ventasdbf.Recordset!fecha = "01/01/01"
         ventasdbf.Recordset.Update
         ventasdbf.Recordset.MoveNext
         Me.LBLREG.Caption = v: Me.LBLREG.Refresh
         v = v + 1
         probar1.Value = v
    Wend
    MsgBox "Proceso Finalizado...", vbInformation
    Me.MousePointer = 0
End If
LBLREG.Caption = 0
End Sub

Private Sub ImpVtaDia()
Dim rs As ADODB.Recordset
Dim PRECOSTO As String
Dim fs As Object
Dim Tdas(1 To 10) As String
Dim cntmp  As ADODB.Connection
On Error Resume Next
MsgBox "LOS ARCHIVOS DE VENTAS DEBERAN ESTAR EN P:\BUZON", vbInformation, "Ubicación de archivos"
Tdas(1) = "VENPER"
Tdas(2) = "VENREF"
Tdas(3) = "VENBRE"
Tdas(4) = "VENCON"
Tdas(5) = "VENCEN"
Tdas(6) = "VENMIN"
Tdas(7) = "VENDOL"
Tdas(8) = "VENSMA"
Tdas(9) = "VENROS"
Tdas(10) = "VENHID"

Set cntmp = New ADODB.Connection
cntmp.ConnectionTimeout = 0
cntmp.CommandTimeout = 0
cntmp.ConnectionString = cCadConex
cntmp.Open
cntmp.Execute "DELETE FROM ventastmp"
Set fs = CreateObject("Scripting.FileSystemObject")

For t = 1 To 10
    sucursal = Mid(Tdas(t), 4, 3)
    cmen = Me.Stb1.SimpleText
    Stb1.SimpleText = Space(50) & "Agregando datos del archivo " & Tdas(t) & " a tabla temporal......"

    cDeli = "|": nreg = 0
    lnuevo = True
    totpza = 0: importe = 0: prec = 0: PREV = 0
    nTur = 0
    cArch = "P:\BUZON\" & Tdas(t) & ".txt"
    If fs.FileExists(cArch) Then
        Open cArch For Input As #1
        Do While Not EOF(1)
            nreg = nreg + 1: LBLREG.Caption = Str(nreg): LBLREG.Refresh
            Line Input #1, lineatexto
            nPos = InStr(1, lineatexto, cDeli)
            CONSEC = Mid(lineatexto, 1, nPos - 1)
            nposant = nPos + 1
            nPos = InStr(nposant, lineatexto, cDeli)
            PRECIO = Mid(lineatexto, nposant, nPos - nposant)
            nposant = nPos + 1
            nPos = InStr(nposant, lineatexto, cDeli)
            cantidad = Mid(lineatexto, nposant, nPos - nposant)
            nposant = nPos + 1
            nPos = InStr(nposant, lineatexto, cDeli)
            fecha = Mid(lineatexto, nposant, nPos - nposant)
            nposant = nPos + 1
            nPos = InStr(nposant, lineatexto, cDeli)
            PRECOSTO = Mid(lineatexto, nposant, nPos - nposant)
            nposant = nPos + 1
            nPos = InStr(nposant, lineatexto, cDeli)
            Preventa = Mid(lineatexto, nposant, nPos - nposant)
            nposant = nPos + 1
            nPos = InStr(nposant, lineatexto, cDeli)
            ESTAC = Mid(lineatexto, nposant, nPos - nposant)
            If lnuevo Then   'Solo se hace en el primer registro
                clave = CONSEC
                FEC = fecha
                EST = ESTAC
                lnuevo = False
            End If
            If sucursal + fecha + ESTAC + CONSEC <> sucursal + FEC + EST + clave Then
                'cntmp.Execute "INSERT INTO Ventastmp VALUES ('" & CONSEC & "'," & PRECIO & "," & CANTIDAD & ",'" & FECHA & "'," & PRECOSTO & "," & PREVENTA & ",'" & ESTAC & "','" & SUCURSAL & "')"
                cntmp.Execute "INSERT INTO Ventastmp(consec,importe,cantidad,fecha,precosto,preventa,estacion,sucursal) VALUES ('" & clave & "'," & importe & "," & totpza & ",'" & FEC & "'," & prec / nTur & "," & PREV / nTur & ",'" & EST & "','" & sucursal & "')"
                totpza = 0: importe = 0: prec = 0: PREV = 0
                nTur = 0
                clave = CONSEC
                FEC = fecha
                totpza = totpza + cantidad
                importe = importe + PRECIO
                prec = prec + PRECOSTO
                PREV = PREV + Preventa
                EST = ESTAC
                nTur = nTur + 1
            Else
                clave = CONSEC
                FEC = fecha
                totpza = totpza + cantidad
                importe = importe + PRECIO
                prec = prec + PRECOSTO
                PREV = PREV + Preventa
                EST = ESTAC
                nTur = nTur + 1
            End If
        Loop
        cntmp.Execute "INSERT INTO Ventastmp(consec,importe,cantidad,fecha,precosto,preventa,estacion,sucursal) VALUES ('" & clave & "'," & importe & "," & totpza & ",'" & FEC & "'," & prec / nTur & "," & PREV / nTur & ",'" & EST & "','" & sucursal & "')"
        Close #1
    End If
Next
Stb1.SimpleText = Space(40) & "Espere un momento procesando ventas de sucursales"
cntmp.Execute "ImportaVentas"
cntmp.Close
Set cntmp = Nothing
MsgBox "El proceso de importación de ventas finalizo correctamente", vbInformation
Stb1.SimpleText = cmen
Stb1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmAreaRecibo.Show
End Sub

Private Sub txtbusca_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Adoventas.Recordset.MoveFirst
    Adoventas.Recordset.Find "PRODUCTO LIKE '" & UCase(txtbusca.Text) & "%'"
End If
End Sub
