VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdespla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desplazamiento de compras"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmTendencias.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame frasug 
      BackColor       =   &H00808080&
      Height          =   2415
      Left            =   3720
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdRanCan 
         Caption         =   "Regresar"
         Height          =   350
         Left            =   2160
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdRanAce 
         Caption         =   "Consultar"
         Height          =   350
         Left            =   600
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   300
         Left            =   2160
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   65077249
         CurrentDate     =   37068
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   300
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   65077249
         CurrentDate     =   37068
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccione el rango de los períodos"
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
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Fecha Final"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Fecha Inicial"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   8280
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   16651
            MinWidth        =   16651
            Text            =   "Para salir sin grabar presione el boton regresar"
            TextSave        =   "Para salir sin grabar presione el boton regresar"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "F4 = Muestra historial"
            TextSave        =   "F4 = Muestra historial"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTend 
      Height          =   375
      Left            =   6000
      Top             =   4680
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11850
      TabIndex        =   0
      Top             =   7545
      Width           =   11910
      Begin VB.CommandButton cmdCamRang 
         Caption         =   "&Cam. Rango"
         Height          =   450
         Left            =   2520
         Picture         =   "frmTendencias.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Actualiza el desplazamiento de productos"
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "&Actualizar"
         Height          =   450
         Left            =   1440
         Picture         =   "frmTendencias.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Actualiza el desplazamiento de productos"
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "Reporte. "
         Height          =   450
         Left            =   360
         Picture         =   "frmTendencias.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Regresar"
         Height          =   450
         Left            =   4680
         Picture         =   "frmTendencias.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   450
         Left            =   3600
         Picture         =   "frmTendencias.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Agrega al detalle de pedidos aquellos productos en los que la columna cantidad solicitada es mayor a cero"
         Top             =   120
         Width           =   1000
      End
      Begin VB.PictureBox cRpt 
         Height          =   480
         Left            =   120
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   22
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCTOS: 0     CAJAS: 0     PIEZAS: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7080
         TabIndex        =   4
         Top             =   120
         Width           =   4575
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdTend 
      Bindings        =   "frmTendencias.frx":0E5C
      Height          =   6855
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12091
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      TabAction       =   2
      RowDividerStyle =   3
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
      Caption         =   "DESPLAZAMIENTO DE COMPRAS REALIZADAS EN LAS ULTIMAS TRES SEMANAS"
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "consec"
         Caption         =   "    CLAVE"
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
         DataField       =   "descripc"
         Caption         =   "          DESCRIPCION DEL PRODUCTO"
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
         Caption         =   "       MEDIDA"
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
         DataField       =   "cantidad3"
         Caption         =   "PRIPER"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "cantidad2"
         Caption         =   "MEDPER"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "cantidad1"
         Caption         =   "ULTPER."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "promedio"
         Caption         =   "PROM. CAJAS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "incant"
         Caption         =   "INV.INI.CAJAS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "incantp"
         Caption         =   " EXIST."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "ininicial"
         Caption         =   "CAJ SUCURSAL"
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
      BeginProperty Column10 
         DataField       =   "maximo"
         Caption         =   "MAXIMO"
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
      BeginProperty Column11 
         DataField       =   "promocion"
         Caption         =   "PROMO CION"
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
      BeginProperty Column12 
         DataField       =   "InStock"
         Caption         =   "STOCK PZAS."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "total"
         Caption         =   "SUGxCAJ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "cantsol"
         Caption         =   "SOLx.CAJ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "cantsolp"
         Caption         =   "SOLxPZA"
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
      SplitCount      =   2
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   1
         Size            =   410
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   4770.142
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
      BeginProperty Split1 
         MarqueeStyle    =   2
         RecordSelectors =   0   'False
         Size            =   250
         BeginProperty Column00 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            ColumnWidth     =   854.929
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraHist 
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   11700
      Begin VB.CommandButton cmdHisReg 
         Cancel          =   -1  'True
         Caption         =   "&Regresar"
         Height          =   450
         Left            =   9960
         Picture         =   "frmTendencias.frx":0E72
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtObserva 
         BackColor       =   &H80000000&
         Height          =   495
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   240
         Width           =   8775
      End
      Begin MSAdodcLib.Adodc AdoEntSur 
         Height          =   330
         Left            =   360
         Top             =   2040
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
         Caption         =   "AdoEntSur"
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
      Begin MSDataGridLib.DataGrid DatgEntSurt 
         Bindings        =   "frmTendencias.frx":0FE4
         Height          =   3735
         Left            =   360
         TabIndex        =   17
         Top             =   4440
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         RowDividerStyle =   0
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
         Caption         =   "ENTRADAS RECIBIDAS [ PEDIDOS ]"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "pp_pedido"
            Caption         =   "     PEDIDO"
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
            DataField       =   "pp_sucursal"
            Caption         =   "SUCURSAL"
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
            DataField       =   "pp_fecCONFIRMA"
            Caption         =   "       FECHA. CONF."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/mm/yy hh:mm AM/PM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "pp_fecrecibe"
            Caption         =   "      FECHA REC."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/mm/yy hh:mm AM/PM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "pp_proveedor"
            Caption         =   "PROV."
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
            DataField       =   "dg_cantsol"
            Caption         =   "CAJ. SOL."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "dg_cantreal"
            Caption         =   "CAJ.REC."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "dg_cantsolp"
            Caption         =   "PZAS.SOL."
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
         BeginProperty Column08 
            DataField       =   "dg_cantrealp"
            Caption         =   "PZAS. REC."
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
            DataField       =   "pp_pedind"
            Caption         =   "PED.IND."
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
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1920.189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   734.74
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoEntPend 
         Height          =   330
         Left            =   360
         Top             =   1560
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
         Caption         =   "AdoEntPend"
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
      Begin MSDataGridLib.DataGrid DatgEntPend 
         Bindings        =   "frmTendencias.frx":0FFC
         Height          =   3615
         Left            =   360
         TabIndex        =   20
         Top             =   720
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         RowDividerStyle =   0
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
         Caption         =   "ENTRADAS PENDIENTES    [ PEDIDOS ]"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "pp_pedido"
            Caption         =   "     PEDIDO"
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
            DataField       =   "pp_sucursal"
            Caption         =   "SUCURSAL"
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
            DataField       =   "pp_fechagen"
            Caption         =   "       FECHA. ELAB."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/mm/yy hh:mm AM/PM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "pp_fecConfirma"
            Caption         =   "      FECHA CONF."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/mm/yy hh:mm AM/PM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "pp_proveedor"
            Caption         =   "PROV."
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
            DataField       =   "dg_cantsol"
            Caption         =   "CAJ. SOL."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "dg_cantsolp"
            Caption         =   "PZAS.SOL."
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
            DataField       =   "dg_promocion"
            Caption         =   "PROM.PACTADA"
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
         BeginProperty Column08 
            DataField       =   "pp_pedind"
            Caption         =   "PED.IND."
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
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   2160
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1920.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   750.047
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmdespla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Tabla As String

Private Sub cmdCamRang_Click()
  Me.frasug.Visible = True
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub cmdGrabar_Click()
'On Error GoTo Error
'Borro todo lo que exista en detalle de facturas
If Not lpprov Then   'Pedido sugerido
    cn.Execute "DELETE FROM [Detallefactura] WHERE Detallefactura.df_pedido = '" & frmCaptPed.txtcampos(0).Text & "' AND df_sugerido = " & Forma
    cn.Execute "INSERT INTO detallefactura( df_prod,df_pedido,df_cantidad,df_cantsol,df_cantsolp,df_sugerido ) SELECT CONSEC, Pedido = '" & frmCaptPed.txtcampos(0).Text & "', Total, CantSol, CantSolp, tipPed = " & Forma & " FROM " & Tabla & " WHERE CANTSOL > 0 OR CANTSOLP > 0"
    cn.Execute "UPDATE DetalleFactura SET DetalleFactura.DF_COSTO = DESCPROD.COSTO FROM DESCPROD, DetalleFactura WHERE DESCPROD.PRODUCTO = DetalleFactura.Df_prod AND DetalleFactura.DF_PEDIDO = '" & frmCaptPed.txtcampos(0).Text & "' AND df_sugerido = " & Forma
    frmCaptPed.AdoDetPed.Refresh
    frmCaptPed.dbgrdDetPed.Refresh
    frmCaptPed.dbgrdDetPed.Visible = True
    frmCaptPed.dbgrdDetPed.Columns(0).Button = True
    frmCaptPed.cmdReporte.Enabled = True
    frmCaptPed.AdoDetPed.Refresh
    frmCaptPed.dbgrdDetPed.Visible = True
    frmCaptPed.dbgrdDetPed.Refresh
    frmCaptPed.dbgrdDetPed.Columns(0).Button = True
    frmCaptPed.cmdReporte.Enabled = True
    frmCaptPed.cmdExporta.Visible = (Trim(frmCaptPed.txtcampos(1).Text) = "ABA")
    poninfo
Else  'Pedidos indirectos
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT * FROM PEDPROVE WHERE pp_pedido = '" & strcveprod & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rsttemp.RecordCount > 0 Then
       cn.Execute "DELETE FROM detalleglobal WHERE dg_pedido = '" & strcveprod & "'"
       cn.Execute "INSERT INTO detalleglobal(dg_pedido,dg_producto,dg_cantsol,dg_cantsolp,dg_cantreal,dg_cantrealp) SELECT Ped = '" & strcveprod & "', consec, cantsol, cantsolp, Cantreal =0 , cantrealp = 0 FROM " & Tabla & " WHERE CANTSOL > 0 OR CANTSOLP > 0"
       frmPedProv.AdoPedpro.Refresh
       frmPedProv.dbgrdPedpro.Columns(1).Width = 5560
    Else
       cn.Execute "INSERT INTO pedprove (pp_proveedor, pp_pedido, pp_fechagen, pp_fecconfirma, pp_confirma, pp_recibe, pp_cancelado, pp_enviado, pp_pedind,PP_sucursal) VALUES (" & _
               "'" & Mid(strcveprod, 1, 3) & "', '" & strcveprod & "', '" & date + Time & "', '" & date + Time & "',1,0,0,'1',1,'" & Trim(Mid(Me.dbgrdTend.Caption, 1, 3)) & "')"
       cn.Execute "INSERT INTO detalleglobal(dg_pedido,dg_producto,dg_cantsol,dg_cantsolp,dg_cantreal,dg_cantrealp) SELECT Ped = '" & strcveprod & "', consec, cantsol, cantsolp, Cantreal =0 , cantrealp = 0 FROM " & Tabla & " WHERE CANTSOL > 0 OR CANTSOLP > 0"
    End If
End If
Unload Me
Exit Sub
Error:
  MsgBox Err.Description

End Sub

Private Sub CmdRefresh_Click()
   AdoTend.Refresh
End Sub

Private Sub cmdHisReg_Click()
  FraHist.Visible = False
  dbgrdTend.Visible = True
  dbgrdTend.SetFocus
End Sub

Private Sub cmdRanCan_Click()
  Me.frasug.Visible = False
End Sub

Private Sub cmdRefrescar_Click()
  AdoTend.Refresh
End Sub

Private Sub cmdReporte_Click()
Dim cMenAnt As String
On Error GoTo Error
   cMensaje = stb1.SimpleText
   stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
   stb1.Refresh
   'frmAreaRecibo.CR1.Connect = "DSN=PITICO;UID=" & cusuario & ";PWD=" & cContraseña
   crpt.Connect = cCadConex
   crpt.ReportFileName = App.Path & "\PeDesPla.rpt"
   crpt.DataFiles(0) = Tabla
   crpt.WindowTitle = "Desplazamiento de productos"
   'cRpt.Formulas(0) = "PROVED = 'PROVEEDOR [ " & IIf(nOp <> 3, frmCaptPed.txtCampos(1).Text, "") & " " & IIf(nOp <> 3, frmCaptPed.cmbProved.Text, FrmReport.cmbCombo(0).Text) & " ]'"
   'cRpt.Formulas(1) = "FECELAB = 'FECHA DE ELABORACION DEL PEDIDO:   " & IIf(nOp <> 3, frmCaptPed.txtCampos(4).Text, FrmReport.txtCampos(0).Text) & "'"
   If nOp <> 3 Then frmAreaRecibo.cr1.Formulas(2) = "NUMPED = 'NUMERO DEL PEDIDO:   " & frmCaptPed.txtcampos(0).Text & "'"
   crpt.Action = 1
   frmCaptPed.StbMensajes.SimpleText = cMensaje
   stb1.SimpleText = cMenAnt
   stb1.Refresh
   Exit Sub
Error:
   MsgBox Err.Description
   frmCaptPed.StbMensajes.SimpleText = cMensaje
End Sub


Private Sub DatgEntPend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 27 Then Me.FraHist.Visible = False
End Sub

Private Sub DatgEntSurt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 27 Then Me.FraHist.Visible = False
End Sub

Private Sub dbgrdTend_AfterColUpdate(ByVal ColIndex As Integer)
If Forma = 1 Then
   If UCase(dbgrdTend.Columns(ColIndex).DataField) = "INCANTP" And Val(dbgrdTend.Columns(ColIndex).Text) >= Val(dbgrdTend.Columns(13).Text) Then
      MsgBox "LA CANTIDAD EN PIEZAS NO PUEDE SER MAYOR AL NUMERO DE PAQUETES POR CAJA, " & _
             "AUMENTE EL NUMERO DE CAJAS", vbExclamation
      dbgrdTend.Columns(14).Text = 0
   End If
End If
'SendKeys "{DOWN}"
keybd_event &H28, 0, 0, 0: keybd_event &H28, 0, &H2, 0
End Sub

Private Sub dbgrdTend_AfterUpdate()
On Error Resume Next
Dim rsttemp As ADODB.Recordset
  Set rsttemp = New ADODB.Recordset
  rsttemp.Open "SELECT COUNT(CONSEC) AS TOTPRO, SUM(cantsol) AS TOTCAJ, SUM(CANTSOLP) AS TOTPZA FROM " & Tabla & " WHERE CANTSOL > 0 OR CANTSOLP > 0 ", cn, adOpenKeyset, adLockOptimistic, adCmdText
  lblInfo.Caption = "PRODUCTOS: " & CStr(rsttemp!totpro) & Space(5) & "CAJAS: " & CStr(rsttemp!totcaj) & Space(5) & "PIEZAS: " & CStr(rsttemp!totpza)
End Sub

Private Sub dbgrdTend_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 115 Then
    Me.dbgrdTend.Visible = False
    FraHist.Visible = True
    cMens = Me.stb1.Panels(1).Text
    stb1.Panels(1).Text = "Espere un momento, consultando historial del producto"
    stb1.Refresh
    Me.FraHist.Caption = Me.AdoTend.Recordset!CONSEC & " " & AdoTend.Recordset!descripc & " " & AdoTend.Recordset!medida
    AdoEntPend.ConnectionString = cCadConex
    AdoEntPend.CommandType = adCmdText
    AdoEntPend.RecordSource = "SELECT * FROM Pedprove, detalleglobal WHERE dg_pedido = pp_pedido AND pp_recibe = 0 AND DG_PRODUCTO = '" & AdoTend.Recordset!CONSEC & "' AND dg_cantsol > 0 ORDER BY pp_fecConfirma DESC"
    AdoEntPend.Refresh
    AdoEntSur.CursorType = adOpenStatic
    AdoEntSur.ConnectionString = cCadConex
    AdoEntSur.CommandType = adCmdText
    AdoEntSur.RecordSource = "SELECT * FROM Pedprove, detalleglobal WHERE dg_pedido = pp_pedido AND pp_recibe = 1 AND DG_PRODUCTO = '" & AdoTend.Recordset!CONSEC & "' AND dg_cantsol > 0 ORDER BY pp_fecrecibe DESC"
    AdoEntSur.Refresh
    stb1.Panels(1).Text = cMens
    stb1.Refresh
 End If
End Sub

Private Sub dbgrdTend_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Me.FraHist.Visible = False
End Sub

Private Sub Form_Load()
Dim N As Integer
 On Error GoTo Error:
 Tabla = "TPEDSUG" & Trim(Mid(cCveDesUsu, 1, 4))  'Existe una tabla temporal por cada usuario
 If nOp <> 3 And nOp <> 1 Then  'Si se llama la forma de reportes,ped indirecto no actualizo el desplazamiento
    cn.Execute "UPDATE " & Tabla & " SET CANTSOL = DF_CANTSOL, CANTSOLP = DF_CANTSOLP FROM DETALLEFACTURA WHERE CONSEC = DF_PROD AND DF_PEDIDO = '" & frmCaptPed.txtcampos(0).Text & "' AND df_sugerido = " & Forma
 End If
 'se sugiere la cantidad solicitada
 cn.Execute "UPDATE " & Tabla & " SET total = PROMEDIO - INVENTARIO.INCANT FROM INVENTARIO WHERE consec = inprod AND INVENTARIO.INCANT < PROMEDIO "
 AdoTend.ConnectionString = cCadConex
 AdoTend.CommandType = adCmdText
 AdoTend.RecordSource = "SELECT * FROM " & Tabla & " ORDER BY descripc,medida"
 AdoTend.Refresh
 dtpFecIni.Value = date - 7
 dtpFecFin.Value = date
 'lpprov = False
 Exit Sub
Error:
   MsgBox Err.Description
End Sub

