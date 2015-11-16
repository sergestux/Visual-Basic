VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPedProvB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos por proveedor pendientes de confimar"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   Icon            =   "frmPedProvB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkCampos 
      Caption         =   "&Proteción"
      DataField       =   "pp_protect"
      DataSource      =   "AdoPedProve"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   58
      Top             =   0
      Width           =   1695
   End
   Begin VB.Frame frmobserva 
      BackColor       =   &H80000001&
      Caption         =   "Observaciones de Recibo del Pedido:"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   1080
      TabIndex        =   55
      Top             =   3360
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton Command1 
         Caption         =   "Regres&ar"
         Height          =   495
         Left            =   6240
         TabIndex        =   57
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtobserva 
         DataField       =   "pp_observarec"
         DataSource      =   "AdoPedProve"
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Frame FRMMODI 
      BackColor       =   &H80000000&
      Caption         =   "CONTRASEÑA PARA MODIFICAR PEDIDO"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   3360
      TabIndex        =   52
      Top             =   2760
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtmodi 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   54
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame fraCon 
      BackColor       =   &H80000001&
      Caption         =   "PROPORCIONE CONTRASEÑA"
      Height          =   1695
      Left            =   4560
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   24
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   23
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblContra 
         BackColor       =   &H80000001&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdfactu 
      Bindings        =   "frmPedProvB.frx":030A
      Height          =   735
      Left            =   120
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1296
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "Factura1"
         Caption         =   "FACTURA 1"
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
         DataField       =   "impfac1"
         Caption         =   "IMP. FAC. 1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Factura2"
         Caption         =   "FACTURA 2"
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
         DataField       =   "impfac2"
         Caption         =   "IMP. FAC. 2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "factura3"
         Caption         =   "FACTURA 3"
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
         DataField       =   "impfac3"
         Caption         =   "IMP. FAC. 3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "factura4"
         Caption         =   "FACTURA 4"
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
         DataField       =   "impfac4"
         Caption         =   "IMP. FAC. 4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "factura5"
         Caption         =   "FACTURA 5"
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
         DataField       =   "impfac5"
         Caption         =   "IMP. FAC. 5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "factura6"
         Caption         =   "FACTURA 6"
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
      BeginProperty Column11 
         DataField       =   "impfac6"
         Caption         =   "IMP. FAC. 6"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Factura7"
         Caption         =   "FACTURA 7"
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
      BeginProperty Column13 
         DataField       =   "impfac7"
         Caption         =   "IMP. FAC. 7"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "factura8"
         Caption         =   "FACTURA 8"
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
      BeginProperty Column15 
         DataField       =   "impfac8"
         Caption         =   "IMP. FAC. 8"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "factura9"
         Caption         =   "FACTURA 9"
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
      BeginProperty Column17 
         DataField       =   "impfac9"
         Caption         =   "IMP. FAC. 9"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "factura10"
         Caption         =   "FACTURA10"
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
      BeginProperty Column19 
         DataField       =   "impfac10"
         Caption         =   "IMP.FAC.10"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
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
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column15 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column17 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   975.118
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicBoton 
      Align           =   2  'Align Bottom
      BackColor       =   &H8000000C&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11850
      TabIndex        =   13
      Top             =   7560
      Width           =   11910
      Begin VB.CommandButton cmdCamPre 
         Caption         =   "&Agr. $"
         Height          =   525
         Left            =   2160
         Picture         =   "frmPedProvB.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Agrega producto para cambio de precio con stock"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Exi"
         Height          =   525
         Left            =   360
         Picture         =   "frmPedProvB.frx":0426
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "reporte de Nota de entrada"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdInven 
         Caption         =   "&Invent."
         Height          =   525
         Left            =   6960
         Picture         =   "frmPedProvB.frx":0958
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Ver inventario"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdobrec 
         Caption         =   "Obs."
         Height          =   525
         Left            =   6360
         Picture         =   "frmPedProvB.frx":0AE2
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdsugerido 
         Caption         =   "Sug"
         Height          =   525
         Left            =   5760
         Picture         =   "frmPedProvB.frx":1014
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Act."
         Height          =   525
         Left            =   3360
         Picture         =   "frmPedProvB.frx":1356
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Actualizar pedido"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grab"
         Height          =   525
         Left            =   5160
         Picture         =   "frmPedProvB.frx":1458
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Grabar pedido e incrementar Inventario"
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Reg"
         Height          =   525
         Left            =   7560
         Picture         =   "frmPedProvB.frx":15CA
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Regresar a la pantalla principal de pedidos por proveedor"
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Nota"
         Height          =   525
         Left            =   960
         Picture         =   "frmPedProvB.frx":173C
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "reporte de Nota de entrada"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdCodBarra 
         Caption         =   "&Barra"
         Height          =   525
         Left            =   3960
         Picture         =   "frmPedProvB.frx":1C6E
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Verificar codigo de barras de productos"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "+ &Prod."
         Height          =   525
         Left            =   4560
         Picture         =   "frmPedProvB.frx":1DA4
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Agregar producto al pedido"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdExporta 
         Caption         =   "&Exp."
         Height          =   525
         Left            =   2760
         Picture         =   "frmPedProvB.frx":1E9E
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Exportar pedido a formato DBF"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdRpteMix 
         Caption         =   "&Sug."
         Height          =   525
         Left            =   1560
         Picture         =   "frmPedProvB.frx":1FA0
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Reporte de pedidos sugeridos a surtirse en bodega"
         Top             =   90
         Width           =   555
      End
      Begin VB.PictureBox CR1 
         Height          =   480
         Left            =   11520
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   60
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Bckorder"
         Height          =   525
         Left            =   0
         TabIndex        =   59
         Top             =   90
         Width           =   375
      End
      Begin VB.Label lblCajas 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PROD:  XX   CAJAS: XX   PIEZAS:  XX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   3240
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDbf 
      Height          =   330
      Left            =   720
      Top             =   5040
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
      Caption         =   "AdoDbf"
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
   Begin VB.Frame fraAvance 
      Caption         =   "Cargando especificaciones de productos "
      Height          =   855
      Left            =   4200
      TabIndex        =   32
      Top             =   3240
      Visible         =   0   'False
      Width           =   4935
      Begin ComctlLib.ProgressBar PGB 
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame fraAgrega 
      BackColor       =   &H80000001&
      Height          =   1695
      Left            =   2280
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox txtPzaSol 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         TabIndex        =   29
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtCajSol 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdCanpro 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7560
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrapro 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   6120
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbprod 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   8775
      End
      Begin VB.Label lblEtiquetas 
         BackColor       =   &H80000001&
         Caption         =   "Piezas solicitadas"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   34
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblEtiquetas 
         BackColor       =   &H80000001&
         Caption         =   "Cajas solicitadas"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "pp_observa"
      DataSource      =   "AdoPedProve"
      Height          =   615
      Index           =   5
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   1320
      Width           =   8895
   End
   Begin MSAdodcLib.Adodc AdoInventario 
      Height          =   330
      Left            =   9720
      Top             =   6000
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
      Caption         =   "AdoInventario"
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
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   7200
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "AdoFacturas"
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
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "pp_fecrecibe"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   0
      Left            =   9480
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "pp_notent"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   3
      Left            =   9480
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkCampos 
      Caption         =   "Re&cibido"
      DataField       =   "pp_recibe"
      DataSource      =   "AdoPedProve"
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
      Index           =   1
      Left            =   7320
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc AdoDetGlo 
      Height          =   330
      Left            =   9720
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoDetGlo"
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
   Begin MSAdodcLib.Adodc AdoPedProve 
      Height          =   330
      Left            =   7200
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoPedProVed"
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
   Begin MSAdodcLib.Adodc AdoProv 
      Height          =   330
      Left            =   4920
      Top             =   5640
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
      Caption         =   "AdoProv"
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
   Begin VB.ComboBox cmbPerCon 
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox cmbProv 
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "pp_perconfirma"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCampos 
      DataField       =   "pp_fechagen"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "pp_proveedor"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "pp_pedido"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCampos 
      Alignment       =   2  'Center
      DataField       =   "pp_fecconfirma"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   4
      Left            =   9480
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkCampos 
      Caption         =   "&Confirmado"
      DataField       =   "pp_confirma"
      DataSource      =   "AdoPedProve"
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
      Index           =   0
      Left            =   7320
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoPedSol 
      Height          =   330
      Left            =   2640
      Top             =   5640
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
      Caption         =   "AdoPedSol"
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
   Begin MSAdodcLib.Adodc AdoPedpro 
      Height          =   330
      Left            =   240
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "AdoPedPro"
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
   Begin MSDataGridLib.DataGrid dbgrdPedsol 
      Bindings        =   "frmPedProvB.frx":24D2
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2990
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
      Caption         =   "PEDIDOS QUE FUNDAMENTAN EL PEDIO POR PROVEEDOR"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "folio"
         Caption         =   "FOLIO PED."
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
         DataField       =   "sucursal"
         Caption         =   "                            SUCURSAL"
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
         DataField       =   "Fecha_sol"
         Caption         =   "         FECHA ELAB."
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
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dbgrdPedpro 
      Bindings        =   "frmPedProvB.frx":24EA
      Height          =   4095
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1.5
      RowHeight       =   15
      TabAction       =   2
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "dg_producto"
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
         DataField       =   "descripc"
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
         DataField       =   "MEDIDA"
         Caption         =   "PRESENTACION"
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
         DataField       =   "DG_CANTSOL"
         Caption         =   "CAJAS SOL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "DG_CANTSOLP"
         Caption         =   "PZAS.SOL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "dg_promocion"
         Caption         =   "PROM.SOL."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "dg_cantreal"
         Caption         =   "CAJ.REC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "DG_CANTREALP"
         Caption         =   "PZAS.REC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "DG_PROMOCIONR"
         Caption         =   "PROM.REC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "activo"
         Caption         =   "Situacion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
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
         Size            =   311
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   4694.74
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   494.929
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar Stbmensajes 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   46
      Top             =   8295
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "                                                                            Para salir presione la tecla [ Esc ]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblactivo 
      Caption         =   "."
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
      Left            =   4080
      TabIndex        =   44
      Top             =   7080
      Width           =   3255
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Nota de entrada"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Persona que confirma"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Fecha de elaboracion"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblEtiquetas 
      Caption         =   "Clave del pedido global"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Menu mped 
      Caption         =   "Pedidos"
      Begin VB.Menu mgrab 
         Caption         =   "Grabar"
      End
      Begin VB.Menu mpro 
         Caption         =   "Prorratea"
      End
      Begin VB.Menu mobse 
         Caption         =   "Observaciones"
      End
      Begin VB.Menu mexp 
         Caption         =   "Exporta"
      End
      Begin VB.Menu minve 
         Caption         =   "Inventario"
      End
      Begin VB.Menu magre 
         Caption         =   "Agregar Producto"
      End
   End
   Begin VB.Menu mbarr 
      Caption         =   "Consulta Codigos de Barras"
   End
   Begin VB.Menu mrep 
      Caption         =   "Reportes"
      Begin VB.Menu mpedsug 
         Caption         =   "Pedidos Sugeridos"
      End
      Begin VB.Menu mnot 
         Caption         =   "Nota de Entrada"
      End
      Begin VB.Menu mexi 
         Caption         =   "Existencia"
      End
   End
   Begin VB.Menu mact 
      Caption         =   "Actualizar"
   End
   Begin VB.Menu msal 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmPedProvB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lModPed As Boolean
Private lMove As Boolean

Private Sub AdoPedpro_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If AdoPedpro.Recordset!activo Then
      lblactivo.Caption = "PRODUCTO ACTIVO"
      lblactivo.ForeColor = QBColor(1)
   Else
      lblactivo.Caption = "PRODUCTO DADO DE BAJA"
      lblactivo.ForeColor = QBColor(4)
End If
lblactivo.Refresh
 
If Not lMove Then Exit Sub
rstDesPro.MoveFirst
rstDesPro.Find "CONSEC = '" & AdoPedpro.Recordset!clave & "'"
If rstDesPro.EOF Then
   lblDespro.Caption = ""
Else
   lblDespro.Caption = Trim(AdoPedpro.Recordset!descripcion) & Chr(13) & " " & Trim(AdoPedpro.Recordset!Present) & "  " & "PROMOCION: " & CStr(rstDesPro!cajas) & "/" & CStr(rstDesPro!encajas)
   If adopredpro.Recordset!activo Then
      lblactivo.Caption = "PRODUCTO ACTIVO"
      lblactivo.ForeColor = "&H80000012&"
   Else
      lblactivo.Caption = "PRODUCTO DADO DE BAJA"
      lblactivo.ForeColor = "&H8000000D&"
   End If
End If
End Sub

Private Sub cmdCamPre_Click()
Dim rs As ADODB.Recordset
  On Error GoTo Error
   Set rs = New ADODB.Recordset
   rs.Open "SELECT incant FROM inventario WHERE inprod = '" & Me.AdoPedpro.Recordset!DG_PRODUCTO & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
   RESP = InputBox("El producto " & AdoPedpro.Recordset!descripc & "  " & AdoPedpro.Recordset!medida & "Actualmente tiene " & rs!InCant & "cajas en existencia." & Chr(13) & Chr(13) & "Proporcione la existencia cuando el sistema informará para el cambio de precio")
   If IsNumeric(RESP) Then
      cn.Execute "INSERT INTO cambpre(producto,invactual,invcamb) VALUES ('" & AdoPedpro.Recordset!DG_PRODUCTO & "'," & rs!InCant & "," & RESP & ")"
      MsgBox "EL PRODUCTO " & AdoPedpro.Recordset!descripc & "  " & AdoPedpro.Recordset!medida & ", SE REGISTRO CORRECTAMENTE PARA CAMBIO DE PRECIO CUANDO LA EXISTENCIA SEA MENOR O IGUAL A " & RESP & " CAJAS", vbInformation
   End If
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdInven_Click()
Dim cmensa As String
  nOp = 31  'se llama al inventario desde traslados
  cmensa = StbMensajes.SimpleText
  StbMensajes.SimpleText = Space(50) & "Espere un momento obteniendo existencia de productos........"
  StbMensajes.Refresh
  frmModInv.Show 1
  StbMensajes.SimpleText = cmensa
  StbMensajes.Refresh

End Sub

Private Sub Command1_Click()

frmobserva.Enabled = False
frmobserva.Visible = False
Me.cmdsugerido.SetFocus
Me.AdoPedProve.Recordset.Update
End Sub

Private Sub cmdobrec_Click()
    frmobserva.Enabled = True
    frmobserva.Visible = True
    txtobserva.SetFocus
End Sub

Private Sub cmdsugerido_Click()
'MsgBox "proceso en contruccion para pedidos de volumen"
'Exit Sub
If AdoPedProve.Recordset!pp_sugeridos Then
   RESP = MsgBox("Ya se Prorrateo el pedido y sus Sugeridos", vbInformation, "SUGERI2")
   'If resp = vbYes Then
      'GENERAENVIOS (AdoPedProve.Recordset!pp_pedido, frecuencia)
      Exit Sub
   'End If
Else
   If AdoPedProve.Recordset!pp_recibe Then
        cn.Close
        cn.CommandTimeout = 0
        cn.ConnectionTimeout = 0
        cn.Open
        ActSugnew (AdoPedProve.Recordset!pp_pedido)
       'ActualizaSug AdoPedProve.Recordset!pp_pedido
   Else
       MsgBox "Es necesario grabar primero el pedido, antes de generar los Traslados...", vbInformation
   End If
End If
End Sub

Private Sub chkCampos_Click(Index As Integer)
If Index = 0 Then
   'If Not AdoPedProve.Recordset.BOF And AdoPedProve.Recordset.EOF Then
   If chkCampos(Index).Visible = True Then
      'MsgBox AdoPedProve.Recordset!pp_confirma = 0
      If AdoPedProve.Recordset!pp_confirma = 0 Then
         txtcampos(4).Text = date + Time
         txtcampos(4).Enabled = False
      End If
         txtcampos(4).Visible = chkCampos(Index).Value = 1
   End If
ElseIf Index = 1 Then
    If Not (AdoPedProve.Recordset.BOF And AdoPedProve.Recordset.EOF) Then
      If AdoPedProve.Recordset!pp_recibe = 0 And chkCampos(Index).Visible Then
         txtRecib(0).Text = date + Time
         txtRecib(0).Enabled = False
      Else
        lblRec(3).Visible = True
      End If
      'For n = 0 To 3
      '    If n > 0 Then lblRec(n).Visible = chkCampos(Index).Value = 1
      '    If n < 3 Then txtRecib(n).Visible = chkCampos(Index).Value = 1
      'Next
      txtRecib(0).Visible = chkCampos(Index).Value = 1
      Me.dbgrdfactu.Visible = chkCampos(Index).Value = 1
   End If
End If
End Sub

Private Sub cmbProv_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub cmbProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{Tab}"
End If

End Sub

Private Sub cmbProv_Validate(Cancel As Boolean)
On Error GoTo Error:
Dim N As Integer
If nOp = 1 Then
   If cmbProv.Text = "" Or IsNull(cmbProv.Text) Then
       MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
       cmbProv.SetFocus
       Cancel = True
   Else
        AdoProv.Recordset.MoveFirst
        AdoProv.Recordset.Find "NomProve = '" & cmbProv.Text & "'"
        If AdoProv.Recordset.EOF = True Then
           MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
           cmbProv.SetFocus
           Cancel = True
        Else
        txtcampos(1).Text = AdoProv.Recordset!prove
        End If
   End If
End If
Exit Sub
Error:
End Sub

Private Sub cmdAgregar_Click()
  cmdReporte.Enabled = False
  CmdAgregar.Enabled = False
  cmdGrabar.Enabled = False
  'Me.cmdsugerido.Enabled = False
  cmdCodBarra.Enabled = False
  cmdregresar.Enabled = False
  fraCon.Visible = True
  txtContra.Text = ""
  txtContra.SetFocus
End Sub

Private Sub cmdCanpro_Click()
  fraAgrega.Visible = False
  cmdReporte.Enabled = True
  CmdAgregar.Enabled = True
  cmdGrabar.Enabled = True
  'cmdsugerido.Enabled = True
  cmdCodBarra.Enabled = True
  cmdregresar.Enabled = True
End Sub

Private Sub cmdCodBarra_Click()
  nOp = 0  'Para que la forma de lectura de codigos de barra sepa de donde se esta llamando
  'frmCodBarrCap.lblEtiquetas(2).Caption = dbgrdPedpro.Columns(1).Text + " " + dbgrdPedpro.Columns(2).Text
  frmCodBarrCap.Show 1
End Sub

Private Sub cmdConAceptar_Click()
Dim rsttemp As ADODB.Recordset
Dim nreg As Integer
If txtContra.Text <> "PITICO00" Then
   MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbExclamation
   txtContra.SetFocus
Else
   If Not lModPed Then
        fraCon.Visible = False
        Set rsttemp = New ADODB.Recordset
        rsttemp.Open "SELECT Descripc, LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, consec FROM TFPRODUC WHERE Claprove = '" & txtcampos(1).Text & "' ORDER BY Descripc", cn, adOpenKeyset, adLockOptimistic, adCmdText
        PGB.Min = 0: nreg = 0
        PGB.Max = rsttemp.RecordCount
        cmbprod.Clear
        fraAvance.Visible = True
        Me.Refresh
        While Not rsttemp.EOF
            nreg = nreg + 1
            PGB.Value = nreg
            cmbprod.AddItem rsttemp!descripc + "  " + rsttemp!medida + "  " + rsttemp!CONSEC
            rsttemp.MoveNext
        Wend
        fraAvance.Visible = False
        fraAgrega.Visible = True
   Else
       Me.fraCon.Visible = False
       dbgrdPedpro.AllowUpdate = True
   End If
End If
End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
  cmdReporte.Enabled = True
  CmdAgregar.Enabled = True
  cmdGrabar.Enabled = True
  'cmdsugerido.Enabled = True
  cmdCodBarra.Enabled = True
  cmdregresar.Enabled = True
End Sub

'Exporta pedidos DETALLE DE LA NOTA DE ENTRADA a dbf
'Con los pedidos recibidos en carbonera para enviarlos a Oficinas centrales
Private Sub CmdExporta_Click()
On Error GoTo Error:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cnFoxPro As ADODB.Connection
   cMenAnt = StbMensajes.SimpleText
   Cmdlg.DialogTitle = "Grabar archivo para enviar pedidos a Oficinas centrales"
   Cmdlg.Filter = "Archivos Visual Fox pro (*.dbf) | *.dbf"
   Cmdlg.CancelError = True
   Cmdlg.ShowSave
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   StbMensajes.SimpleText = Space(45) & "Grabando archivo " & cRutArc
   StbMensajes.Refresh
   
   For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
   Next
   cruta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

   If MsgBox("DESEAS LIMPIAR EL ARCHIVO A ENVIAR", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        StbMensajes.SimpleText = Space(65) & "Limpiando archivo " & cArch
        StbMensajes.Refresh
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFile("P:\ESPEDCAR.DBF")
        f.Copy cRutArc, True
   End If

   Set rsttemp = New ADODB.Recordset
   AdoDbf.CommandType = adCmdText
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch
   AdoDbf.Refresh
     
   rsttemp.Open "SELECT * FROM DetalleNota WHERE ClaveNota = '" & txtRecib(3).Text & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
   While Not rsttemp.EOF
      StbMensajes.SimpleText = Space(75) & "Exportando producto con la clave: " & CStr(rsttemp!producto)
      StbMensajes.Refresh
      AdoDbf.Recordset.AddNew
      AdoDbf.Recordset!Clavenota = rsttemp!Clavenota
      AdoDbf.Recordset!producto = rsttemp!producto
      AdoDbf.Recordset!CantSolc = rsttemp!cantsol
      AdoDbf.Recordset!cantsolp = rsttemp!cantsolp
      AdoDbf.Recordset!CantRecC = rsttemp!cantrec
      AdoDbf.Recordset!cantrecp = rsttemp!cantrecp
      AdoDbf.Recordset!costo = rsttemp!costo
      AdoDbf.Recordset!ImpFac = Trim(txtRecib(1).Text)
      AdoDbf.Recordset!sucursal = Trim(Mid(cSucursal, 1, 3))
      AdoDbf.Recordset!FecRec = txtRecib(0).Text
      AdoDbf.Recordset!Importado = False
      AdoDbf.Recordset.Update
      rsttemp.MoveNext
   Wend
   AdoDbf.Recordset.Close
   StbMensajes.SimpleText = cMenAnt
  Exit Sub
Error:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
   End If
   StbMensajes.SimpleText = cMenAnt

End Sub

Private Function VERIFICA() As Boolean
StbMensajes.SimpleText = Space(25) & "Verificando Cantidades Recibidas..."
StbMensajes.Refresh
AdoPedpro.Recordset.MoveFirst
total = 0
While Not AdoPedpro.Recordset.EOF
     total = total + AdoPedpro.Recordset!dg_cantreal
     AdoPedpro.Recordset.MoveNext
Wend
If Not total >= 1 Then
   VERIFICA = False
Else
   VERIFICA = True
End If
End Function


Private Function siback() As Boolean
AdoPedpro.Recordset.MoveFirst
total = 0
totped = 0
While Not AdoPedpro.Recordset.EOF
     total = total + AdoPedpro.Recordset!dg_cantreal
     totped = totped + AdoPedpro.Recordset!dg_cantsol
     AdoPedpro.Recordset.MoveNext
Wend
If total < totped Then
   siback = True
Else
   siback = False
End If
End Function

Public Sub listaexistencias(clave As String)
Dim rsttemp As ADODB.Recordset
PROVEEDOR = clave 'Me.AdoPedProve.Recordset!pp_proveedor
Set rsttemp = New ADODB.Recordset
rsttemp.Open "SELECT inprod,incant,incantpza, descripc,contenid,medida,paquetes FROM tfproduc,INVENTARIO WHERE consec = inprod and claprove  = '" & PROVEEDOR & "' order by descripc ", cn, adOpenDynamic, adLockOptimistic, adCmdText
Open App.Path & "\EXISTENCIA.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
Print #1, "     EXISTENCIAS DE " & PROVEEDOR & " DEL   "; UCase(Format(date, "long date"))
Print #1,   ' Imprime una línea en blanco en el archivo.
Print #1, "=========================================================================================="
Print #1, "          PRODUCTO                                                    CAJAS        PIEZAS"
Print #1, "=========================================================================================="
While Not rsttemp.EOF
     CAD = Mid(rsttemp!Inprod, 1, 10) & "  " & Mid(rsttemp!descripc, 1, 20) & "  " & Format(Trim(rsttemp!PAQUETES), "000") & " x " & Format(Trim(rsttemp!CONTENID), "000.000") & " " & Mid(rsttemp!medida, 1, 4)
     Print #1, Mid(CAD, 1, 100);
     'If Len(cad) < 45 Then
     '     Print #1, vbTab;
     'End If
     Print #1, vbTab & vbTab & vbTab;
     Print #1, rsttemp!InCant & vbTab;
     Print #1, rsttemp!InCantPza
     rsttemp.MoveNext
Wend
Close #1
Set rsttemp = Nothing
Handle = Shell("NOTEPAD " & App.Path & "\EXISTENCIA.TXT", 1)
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo Error
Dim FecConf As Date
Dim lback As Boolean
Dim rsttemp As ADODB.Recordset
Dim rstProPed As ADODB.Recordset
Dim rstInvent As ADODB.Recordset
Dim lTrans As Boolean
  lTrans = False  'Flag de la Transaccion para saber cuando empieza
  lMove = False
  'Cuando se confirma
  'If chkCampos(0).Value = 0 Then
  '   MsgBox "Es necesario activar la casilla de pedido confirmado", vbExclamation
  '   chkCampos(0).SetFocus
  '   Exit Sub
  'Cuando se recibe
  If chkCampos(0).Value = 1 And chkCampos(1).Value = 0 And chkCampos(1).Visible = True Then
     MsgBox "Es necesario activar la casilla de pedido recibido", vbExclamation
     chkCampos(1).SetFocus
     Exit Sub
  End If
  If chkCampos(1).Value = 1 And chkCampos(1).Visible = True Then
     If Not VERIFICA Then
        MsgBox "Se ha detectado un recibo en ceros, Se tomara de Existencias de Bodega ... ", vbCritical, "CEROS"
        'Exit Sub
     End If
  End If
  'si se confirma el pedido
  If chkCampos(0).Value = 1 And chkCampos(1).Value = 0 Then
     'Grabo los datos generales del pedido por proveedor
     AdoPedProve.Recordset!pp_recibe = 0
     AdoPedProve.Recordset.Update
     'Agrego el detalle del pedido por proveedor para prepararlo para recibirlo
     Set rsttemp = New ADODB.Recordset
     rsttemp.ActiveConnection = cCadConex
     rsttemp.CursorType = adOpenKeyset
     rsttemp.LockType = adLockOptimistic
     rsttemp.Source = "DETALLEGLOBAL"
     rsttemp.Open
     AdoPedpro.Recordset.MoveFirst
     While Not AdoPedpro.Recordset.EOF
         rsttemp.AddNew
         rsttemp!dg_pedido = txtcampos(0).Text
         rsttemp!DG_PRODUCTO = AdoPedpro.Recordset!clave
         rsttemp!DG_CANTIDAD = AdoPedpro.Recordset!TotSol
         rsttemp!dg_cantrec = 0
         rsttemp.Update
         AdoPedpro.Recordset.MoveNext
     Wend
     'Confirmo los pedidos por tienda que incluye el Ped por proveedor
     rsttemp.Close
     rsttemp.Source = "SELECT * FROM PEDIDOS WHERE P_proveedor = '" & txtcampos(1).Text & "'  AND p_situacion = 0"
     rsttemp.Open
     While Not rsttemp.EOF
        rsttemp!p_fecConfirma = txtcampos(4).Text
        rsttemp!p_situacion = 1
        rsttemp!P_pedproveedor = txtcampos(0).Text
        rsttemp.Update
        rsttemp.MoveNext
     Wend
  'Si recibe el pedido
  ElseIf chkCampos(1).Value = 1 Then
  'ElseIf chkCampos(0).Value = 1 And chkCampos(1).Value = 1 Then
     RESP = MsgBox("Desea listar las Existencias antes de Generar el Proceso ", vbYesNo)
     If RESP = vbYes Then listaexistencias (Me.AdoPedProve.Recordset!pp_proveedor)
     men1 = "Este proceso Incrementa las Existencias de acuerdo a lo recibido" & vbCrLf & "Deseas Continuar ? " & vbCr
     RESP = MsgBox(men1, vbYesNo, "AUMENTAR INVENTARIO")
     If RESP = vbNo Then
        Exit Sub
     End If
     Set rsttemp = New ADODB.Recordset
     cn.Execute "UPDATE DETALLEGLOBAL SET DG_COSTO =0 WHERE DG_PEDIDO =  '" & txtcampos(0).Text & "' AND DG_COSTO IS NULL"
     cn.BeginTrans: lTrans = True
     FolNot = "N" + txtcampos(0).Text
     StbMensajes.SimpleText = Space(25) & "Cargando inventario para su actualizacion"
     StbMensajes.Refresh
     Set rstInvent = New ADODB.Recordset
     Set rstProPed = New ADODB.Recordset
     dbgrdPedpro.AllowUpdate = False
     AdoPedpro.Recordset.MoveFirst
     While Not AdoPedpro.Recordset.EOF
         'Afecto inventario
         rstInvent.Open "SELECT inprod,incant,incantpza FROM INVENTARIO WHERE Inprod = '" & AdoPedpro.Recordset!DG_PRODUCTO & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
         If rstInvent.BOF And rstInvent.EOF Then
             MsgBox "NO EXISTE EN EL INVENTARIO EL ARTICULO " & Chr(13) & _
             AdoPedpro.Recordset!DG_PRODUCTO & "  " & rstProPed!descripc & Chr(13) & CStr(rstProPed!PAQUETES) & " X " & CStr(rstProPed!CONTENID) & " " & rstProPed!medida & Chr(13) & _
             "A CONTINUACION SE DARA DE ALTA EN EL INVENTARIO", vbInformation
             rstInvent.AddNew
             rstInvent!Inprod = AdoPedpro.Recordset!DG_PRODUCTO
             rstInvent!Insucursal = "3"
             rstInvent!InObserva = " "
             rstInvent!infeccaduprox = "1/1/1900"
             rstInvent!InInicial = 0
             rstInvent!instock = 0
         End If
        exiant = rstInvent!InCant
        'Sumo cantidad recibida en Cajas, piezas y promociones
        rstInvent!InCant = rstInvent!InCant + AdoPedpro.Recordset!dg_cantreal + IIf(IsNull(AdoPedpro.Recordset!DG_PROMOCIONR), 0, AdoPedpro.Recordset!DG_PROMOCIONR)
        rstInvent.Update
        rstInvent.Close
        cn.Execute "UPDATE detalleGlobal SET dg_existencia = " & exiant & " WHERE dg_producto = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND dg_pedido = '" & txtcampos(0).Text & "'"
        AdoPedpro.Recordset.Update
        AdoPedpro.Recordset.MoveNext
     Wend
     StbMensajes.SimpleText = Space(25) & "Generando Nota de entrada"
     StbMensajes.Refresh
     'Obtengo los importes de cantidades solicitada y recibida
     Set rsttemp = New ADODB.Recordset
     rsttemp.Open "SELECT SUM(DetalleGlobal.dg_cantsol * detalleglobal.dg_costo) + SUM(DetalleGlobal.dg_cantsolP * (dg_costo / tfProduc.Paquetes )) AS ImptSol, SUM(DetalleGlobal.dg_cantreal * Tfproduc.Precosto) + SUM(DetalleGlobal.dg_cantrealp * (Tfproduc.Precosto / tfProduc.Paquetes  )) AS ImptRec FROM DetalleGlobal,Tfproduc WHERE DetalleGlobal.dg_producto = Tfproduc.Consec AND DetalleGlobal.dg_pedido = '" & txtcampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
     ntot = 0
     'AdoFacturas.Refresh
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac1), 0, AdoFacturas.Recordset!Impfac1)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac2), 0, AdoFacturas.Recordset!Impfac2)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac3), 0, AdoFacturas.Recordset!Impfac3)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac4), 0, AdoFacturas.Recordset!Impfac4)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac5), 0, AdoFacturas.Recordset!Impfac5)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac6), 0, AdoFacturas.Recordset!Impfac6)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac7), 0, AdoFacturas.Recordset!Impfac7)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac8), 0, AdoFacturas.Recordset!Impfac8)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac9), 0, AdoFacturas.Recordset!Impfac9)
     ntot = ntot + IIf(IsNull(AdoFacturas.Recordset!Impfac10), 0, AdoFacturas.Recordset!Impfac10)

     'Agrego Nota de entrada
     cn.Execute "UPDATE NotaEntrada SET ImporteSol = " & IIf(IsNull(rsttemp!ImptSol), 0, rsttemp!ImptSol) & ", ImporteRec = " & rsttemp!ImptRec & ", DifeProd = " & IIf(IsNull(rsttemp!ImptSol), 0, rsttemp!ImptSol) - rsttemp!ImptRec & ", DifPrecio = " & IIf(IsNull(ntot), 0, ntot) - rsttemp!ImptRec & " WHERE pedido = '" & Trim(txtcampos(0).Text) & "'"
     cn.Execute "INSERT INTO [DetalleNota](ClaveNota,Producto,Cantsol,cantsolp,Cantrec,CantrecP,costo) SELECT CveNota = '" & FolNot & "', dg_producto, dg_cantsol, dg_cantsolp, dg_cantreal, dg_cantrealp, dg_costo FROM [DetalleGlobal] WHERE dg_pedido = '" & txtcampos(0).Text & "'"
     AdoPedProve.Recordset!pp_pedido = txtcampos(0).Text
     AdoPedProve.Recordset!pp_proveedor = txtcampos(1).Text
     AdoPedProve.Recordset!PP_FECHAGEN = txtcampos(3).Text
     AdoPedProve.Recordset!pp_recibe = 1
     AdoPedProve.Recordset!pp_fecrecibe = txtRecib(0).Text
     AdoPedProve.Recordset!pp_NotEnt = Trim(FolNot)
     

     'StbMensajes.SimpleText = Space(25) & "Espere un momento, realizando prorrateo de pedidos sugeridos"
     'StbMensajes.Refresh
     AdoPedProve.Recordset.Update
     cmdGrabar.Enabled = False
     cmdGrabar.Enabled = False
     CmdAgregar.Enabled = False
     cmdCodBarra.Enabled = False
     chkCampos(1).Enabled = False
     cmdReporte.Enabled = True
     cmdReporte.Visible = True
     cn.CommitTrans: lTrans = False
  End If
If siback Then ' BACKORDER
   Set rsttemp = New ADODB.Recordset
   rsttemp.Open "SELECT Nomprove, Backorder FROM CATPROV WHERE Prove = '" & Trim(AdoPedProve.Recordset!pp_proveedor) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rsttemp!backorder Then
     cresp = MsgBox("AL PROVEEDOR " & rsttemp!NOMPROVE & Chr(13) & "SE LE PERMITEN VARIAS ENTREGAS." & Chr(13) & Space(15) & "DESEA CREAR BACKORDER?", vbYesNo + vbInformation)
        If cresp = vbYes Then generaback
   End If
End If
cmdsugerido.Enabled = True
cmdsugerido.Visible = True
listaexistencias (Me.AdoPedProve.Recordset!pp_proveedor)
CAD = "Presione el Boton [sug], para realizar el Prorrateo de Sugeridos"
MsgBox "Se genero la nota de entrada  " & FolNot & vbCrLf & CAD, vbExclamation
Exit Sub
Error:
  MsgBox "OCURRIO EL SIGUIENTE ERROR: " + Chr(13) + UCase(Err.Description), vbCritical
  If lTrans = True Then
      MsgBox "A CONTINUACION SE DESHARAN LAS MODIFICACIONES REALIZADAS EN EL INVENTARIO", vbCritical
      cn.RollbackTrans
      AdoPedProve.Recordset!pp_recibe = 0
      AdoPedProve.Recordset.Update
  End If
  Unload Me
End Sub

Private Sub generaback()
Dim rsttemp As ADODB.Recordset
Set rsttemp = New ADODB.Recordset
rsttemp.Source = "SELECT pp_pedBack, MAX(pp_NumBack) AS NumBack FROM Pedprove WHERE pp_pedback = '" & IIf(Mid(txtcampos(0).Text, 1, 1) = "B", AdoPedProve.Recordset!pp_pedback, txtcampos(0).Text) & "' GROUP BY pp_pedback"
rsttemp.Open , cn, adOpenKeyset, adLockOptimistic, adCmdText
If rsttemp.EOF Then
            folback = 1
            FolPed = "B" & CStr(folback) & txtcampos(0).Text
Else
            folback = rsttemp!NumBack + 1
            If IsNull(folback) Then
                folback = 100
            End If
            FolPed = "B" & CStr(folback) & rsttemp!pp_pedback
End If

cn.Execute "INSERT INTO [PedProve](pp_proveedor,pp_pedido,pp_fechagen,pp_confirma,pp_recibe,pp_numback,pp_pedback ) VALUES " & _
    "('" & txtcampos(1).Text & "','" & FolPed & "','" & date + Time & "',1,0," & folback & ",'" & IIf(folback = 1, txtcampos(0).Text, rsttemp!pp_pedback) & "')"
CAD = "INSERT INTO DetalleGlobal(dg_pedido,dg_producto,dg_cantsol,dg_cantsolp,dg_cantreal,dg_costo) " & _
    " SELECT FOLIO = '" & Trim(FolPed) & "', dg_producto,dg_cantsol - dg_cantreal ,dg_cantsolp - dg_cantrealp , dg_cantreal = 0 , dg_costo " & _
    " from detalleglobal where  dg_cantsol > dg_cantreal and  dg_pedido = '" & Trim(txtcampos(0).Text) & "'"

'MsgBox CAD
cn.Execute CAD
Set rsttemp = Nothing
End Sub


Private Sub cmdGrapro_Click()
On Error GoTo Error:
Dim rstExis As ADODB.Recordset
  If Not IsNumeric(txtCajSol.Text) Or Not IsNumeric(txtPzaSol.Text) Then
     MsgBox "LA CANTIDAD EN CAJAS Y PIEZAS DEBE SER NUMERICA", vbExclamation
     Exit Sub
  End If
  cclave = Trim(Mid(cmbprod.Text, Len(Trim(cmbprod.Text)) - 8))
  Set rstExis = New ADODB.Recordset
  rstExis.Open "SELECT * FROM DETALLEGLOBAL WHERE dg_producto = '" & cclave & "' AND dg_pedido = '" & txtcampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
  If rstExis.RecordCount > 0 Then
     MsgBox "EL ARTICULO SELECCIONADO YA EXISTE EN EL PEDIDO", vbExclamation
     Exit Sub
  End If
  If MsgBox("REALMENTE DESEAS AGREGAR EL PRODUCTO " & Chr(13) & cmbprod.Text, vbQuestion + vbYesNo) = vbYes Then
     cn.Execute "INSERT INTO DETALLEGLOBAL(dg_pedido,dg_producto,dg_cantsol,dg_cantsolp) VALUES('" & txtcampos(0).Text & "','" & cclave & "'," & txtCajSol.Text & "," & txtPzaSol.Text & ")"
     AdoPedpro.Refresh
  End If
  Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub CmdRefresh_Click()
  AdoPedpro.Refresh
End Sub

Private Sub cmdRegresar_Click()
Unload Me
End Sub

Private Sub cmdReporte_Click()
'On Error GoTo ERROR:
'cMensaje = stbmensajes.SimpleText
StbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
StbMensajes.Refresh
cr1.Connect = cCadConex
If cModo = "VERCONF" Then
    cr1.ReportFileName = App.Path & "\Pedprove.rpt"
    cr1.WindowTitle = "Pedido por proveedor numero " & txtcampos(0).Text
    cr1.Formulas(0) = "FORMSELEC = '" & txtcampos(0).Text & "'"
    cr1.Formulas(1) = "FECELAB = 'FECHA DE ELABORACION:  " & txtcampos(3).Text & "'"
    cr1.Formulas(2) = "NUMPED = 'NUMERO DE PEDIDO:  " & txtcampos(0).Text & "'"
    cr1.Formulas(3) = "PROVED = 'PROVEEDOR " & txtcampos(1).Text & Space(3) & frmpedBod.cmbProved.Text & "' "
    cr1.Formulas(4) = "FECCONF = 'FECHA DE CONFIRM.:  " & txtcampos(4).Text & "' "
    cr1.Formulas(5) = "FECCONF = 'FECHA DE CONFIRM.:  " & txtcampos(4).Text & "' "
    
Else
    cr1.ReportFileName = App.Path & "\prNotEnt.rpt"
    cr1.WindowTitle = "Nota de entrada del pedido " & txtcampos(0).Text
 '   CR1.Formulas(0) = "FORMSELEC = '" & txtCampos(0).Text & "'"
    cr1.Formulas(0) = "PROVED = 'PROVEEDOR [ " & txtcampos(1).Text & Space(3) & frmpedBod.cmbProved.Text & " ]'"
    cr1.Formulas(1) = "FOLNOTENT = 'FOLIO " & Trim(txtRecib(3).Text) & "'"
    'If IsNull(Me.AdoPedProve.Recordset!pp_observarec) Then
    '    CR1.Formulas(2) = "OBSERVARE = ' SIN OBSERVACIONES EN LAS FACTURAS RELACIONADAS CON ESTE PEDIDO '"
    'Else
    '    CR1.Formulas(2) = "observare = ' ' "
    'End If
    cr1.SQLQuery = " SELECT PEDPROVE.pp_pedido, PEDPROVE.pp_fechagen, PEDPROVE.pp_fecconfirma, PEDPROVE.pp_fecrecibe, " & Chr(13) & _
                            "DETALLEGLOBAL.dg_producto, DETALLEGLOBAL.dg_cantsol, DETALLEGLOBAL.dg_cantsolp, DETALLEGLOBAL.dg_cantreal, DETALLEGLOBAL.dg_cantrealp, DETALLEGLOBAL.dg_promocion, DETALLEGLOBAL.dg_costo, DETALLEGLOBAL.dg_promocionR, " & Chr(13) & _
                            "NOTAENTRADA.factura1, NOTAENTRADA.impfac1, NOTAENTRADA.Factura2, NOTAENTRADA.impfac2, NOTAENTRADA.factura3, NOTAENTRADA.impfac3, NOTAENTRADA.factura4, NOTAENTRADA.impfac4, NOTAENTRADA.factura5, NOTAENTRADA.impfac5, NOTAENTRADA.impfac6, NOTAENTRADA.impfac7, NOTAENTRADA.impfac8, NOTAENTRADA.impfac9, NOTAENTRADA.impfac10, " & Chr(13) & _
                            "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES" & Chr(13) & _
                    " FROM pitico.dbo.TFPRODUC TFPRODUC, " & _
                         "pitico.dbo.DETALLEGLOBAL DETALLEGLOBAL, " & _
                         "pitico.dbo.PEDPROVE PEDPROVE, " & _
                         "pitico.dbo.NOTAENTRADA NOTAENTRADA " & Chr(13) & _
                    " WHERE PEDPROVE.pp_pedido = DETALLEGLOBAL.dg_pedido AND " & _
                         "PEDPROVE.pp_pedido = NOTAENTRADA.pedido AND " & _
                         "DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND " & _
                         "PEDPROVE.pp_pedido = '" & txtcampos(0).Text & "' AND (DETALLEGLOBAL.dg_cantsol > 0 OR DETALLEGLOBAL.dg_cantsolp > 0 ) " & Chr(13) & _
                    " ORDER BY PEDPROVE.pp_pedido ASC, TFPRODUC.DESCRIPC ASC, TFPRODUC.CONTENID ASC"
End If
'MsgBox CR1.SQLQuery
'Debug.Print CR1.SQLQuery
cr1.Action = 1
StbMensajes.SimpleText = cMensaje
StbMensajes.Refresh
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdRpteMix_Click()
On Error GoTo Error:
 cMensaje = StbMensajes.SimpleText
 StbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
 StbMensajes.Refresh
 cr1.Connect = cCadConex
 cARcRpt = App.Path & "\provesug.rpt"
 ccondrpt = "FORMSELEC = {PEDIDOS.p_pedproveedor} = '" & txtcampos(0).Text & "'"
 cr1.WindowTitle = "Sugeridos del pedido por proveedor con folio " & txtcampos(0).Text
 cr1.ReportFileName = cARcRpt
 cr1.SQLQuery = "SELECT PEDIDOS.p_pedido, PEDIDOS.p_proveedor, PEDIDOS.p_fecped, PEDIDOS.p_fecconfirma, PEDIDOS.p_pedproveedor, PEDIDOS.p_traslado, " & _
                        "DETALLEFACTURA.df_cantreal, DETALLEFACTURA.df_cantsol, DETALLEFACTURA.df_cantsolp, " & _
                        "CATPROV.NOMPROVE, CATTIENDA.tidescrip, TFPRODUC.CLAPROVE, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & Chr(13) & _
                "FROM pitico.dbo.PEDIDOS PEDIDOS, " & _
                        "pitico.dbo.DETALLEFACTURA DETALLEFACTURA, " & _
                        "pitico.dbo.CATPROV CATPROV, " & _
                        "pitico.dbo.CATTIENDA CATTIENDA, " & _
                        "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                "WHERE PEDIDOS.p_pedido = DETALLEFACTURA.df_pedido AND " & _
                        "PEDIDOS.p_proveedor = CATPROV.PROVE AND " & _
                        "PEDIDOS.p_sucursal = CATTIENDA.ticlave AND " & _
                        "DETALLEFACTURA.df_prod = TFPRODUC.CONSEC AND " & _
                        "PEDIDOS.p_pedproveedor = '" & Trim(txtcampos(0).Text) & "' AND DETALLEFACTURA.df_sugerido  = 1 " & Chr(13) & _
                "ORDER BY PEDIDOS.p_pedido ASC, TFPRODUC.DESCRIPC ASC, TFPRODUC.CONTENID ASC"
                       '"PEDIDOS.p_pedproveedor = '" & Trim(txtCampos(0).Text) & "' AND TFPRODUC.CLAPROVE = '" & Trim(txtCampos(1).Text) & "' " & Chr(13) & _
 'MsgBox CR1.SQLQuery
 cr1.Formulas(0) = ccondrpt
 cr1.Formulas(1) = "FECELAB = 'FECHA DE ELABORACION:  " & txtcampos(3).Text & "'"
 cr1.Formulas(2) = "NUMPED = 'NUMERO DE PEDIDO:  " & txtcampos(0).Text & "'"
 cr1.Formulas(3) = "FECCONF = 'FECHA DE CONFIRM.:  " & IIf(chkCampos(0).Value, txtcampos(4).Text, "") & "' "
 cr1.Formulas(4) = ""
 cr1.Action = 1
 StbMensajes.SimpleText = cMensaje
 StbMensajes.Refresh
 Exit Sub
Error:
 MsgBox Err.Description
End Sub


Private Sub Command2_Click()
listaexistencias (Me.AdoPedProve.Recordset!pp_proveedor)
Exit Sub
cr1.ReportFileName = App.Path & "\ExAntDes.rpt"
cr1.WindowTitle = "Reporte de entrada al inventario del pedido " & txtcampos(0).Text
cr1.Formulas(0) = "FORMSELEC = {PEDPROVE.pp_pedido} = '" & txtcampos(0).Text & "' AND ( {DETALLEGLOBAL.dg_cantreal} > 0 OR {DETALLEGLOBAL.dg_cantrealp} > 0 )"
cr1.Action = 1
End Sub

Private Sub Command3_Click()
Call generaback
End Sub

Private Sub dbgrdfactu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub dismInuyeinventario(ANTERIOR As Integer, NUEVO As Integer)
dis = ANTERIOR - NUEVO
If dis > 0 Then
   CAD = "update inventario set incant = incant - " & dis & " WHERE INPROD = '" & Trim(AdoPedpro.Recordset!DG_PRODUCTO) & "'"
   'cn.Execute " UPDATE DETALLEGLOBAL SET DG_CANTREAL = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!dg_producto & "' AND DG_PEDIDO = '" & Trim(txtCampos(0).Text) & "'"
Else
   MsgBox "No esta permitido Aumentar de lo capturado ", vbInformation
End If
cn.Execute CAD


End Sub
Private Sub dbgrdPedpro_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim marca
Dim ANTERIOR As Integer
Dim NUEVO As Integer
On Error GoTo Error:
marca = AdoPedpro.Recordset.Bookmark
'If Not AdoPedProve.Recordset!pp_recibe Then
    If UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "DG_CANTREAL" Then
        ANTERIOR = AdoPedpro.Recordset!dg_cantreal
        NUEVO = dbgrdPedpro.Columns(ColIndex).Text
        cn.Execute " UPDATE DETALLEGLOBAL SET DG_CANTREAL = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND DG_PEDIDO = '" & Trim(txtcampos(0).Text) & "'"
        If MODIFICADO Then
           Call dismInuyeinventario(ANTERIOR, NUEVO)
           'ADEMAS SE REGISTRA EL MOVIMIENTO EN LA TABLA DETALLAGLOBAL
           cn.Execute "UPDATE DETALLEGLOBAL SET DG_ANTERIOR = " & ANTERIOR & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND DG_PEDIDO = '" & Trim(txtcampos(0).Text) & "'"
        End If
    ElseIf UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "DG_CANTREALP" Then
        cn.Execute "UPDATE DETALLEGLOBAL SET DG_CANTREALP = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND DG_PEDIDO = '" & Trim(txtcampos(0).Text) & "'"
    ElseIf UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "DG_PROMOCIONR" Then
        cn.Execute "UPDATE DETALLEGLOBAL SET DG_PROMOCIONR = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND DG_PEDIDO = '" & Trim(txtcampos(0).Text) & "'"
    End If
    PonPie  'pone cantidades en cajas y piezas
'Else 'Si ya es recibido
'    nCant = dbgrdPedpro.Columns(6).Text
'   If MsgBox("CONFIRMA SI SE ACTUALIZA EL PEDIDO, INVENTARIO Y NOTA DE ENTRADA", vbQuestion + vbYesNo) = vbYes Then
'    If UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "DG_CANTREAL" Then
'       cn.Execute "UPDATE DETALLEGLOBAL SET DG_CANTREAL = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE DG_PRODUCTO = '" & AdoPedpro.Recordset!dg_producto & "' AND DG_PEDIDO = '" & Trim(txtCampos(0).Text) & "'"
'       cn.Execute "UPDATE inventario SET INCANT = INCANT + " & dbgrdPedpro.Columns(6).Text - AdoPedpro.Recordset!dg_cantreal & " WHERE inprod = '" & AdoPedpro.Recordset!dg_producto & "'"
'       cn.Execute "UPDATE Detallenota SET cantrec = " & nCant & " WHERE producto = '" & AdoPedpro.Recordset!dg_producto & "' AND clavenota = '" & "N" + Trim(txtCampos(0).Text) & "'"
'       MsgBox "ACTUALIZADO...", vbInformation
'    End If
'   End If
'End If
AdoPedpro.Refresh
dbgrdPedpro.Refresh
AdoPedpro.Recordset.Bookmark = marca
Cancel = True
SendKeys "{DOWN}"
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub dbgrdRec_AfterColUpdate(ByVal ColIndex As Integer)
'If ColIndex = 1 And AdoDetGlo.Recordset!dg_cantreal > AdoDetGlo.Recordset!dg_cantsol Then
   'Temporalmente se desactiva mientras se define lo de las promociones SAVE
   'MsgBox "LA CANTIDAD RECIBIDA EN CAJAS NO PUEDE SER MAYOR A LA SOLICITADA", vbExclamation
   'AdoDetGlo.Recordset!dg_cantreal = 0
'ElseIf ColIndex = 2 And AdoDetGlo.Recordset!dg_cantrealP > AdoDetGlo.Recordset!dg_cantsolp Then
   'MsgBox "LA CANTIDAD RECIBIDA EN PIEZAS NO PUEDE SER MAYOR A LA SOLICITADA", vbExclamation
   'AdoDetGlo.Recordset!dg_cantrealP = 0
'End If

End Sub

Private Sub dbgrdRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then
   frmCalc.Show
'If KeyCode = 116 Then
'   lModPed = True
'   fraCon.Visible = True
'   txtContra.SetFocus
'End If
ElseIf KeyCode = 113 Then
   'RUTINA PARA PERMITIR MODIFCAR EL PEDIDO
   FRMMODI.Enabled = True
   FRMMODI.Visible = True
   txtmodi.Text = ""
   txtmodi.SetFocus
End If
End Sub

Private Sub MODI_PEDIDO()
If Trim(txtmodi.Text) = "098ZX" Then
   MsgBox UCase("Recuerde que solo se permite Modificar las cantidades de Recibo ") & Chr(13) & UCase("  Pero No se Modifica el Inventario ") & Chr(13) & UCase(" Y TAMPOCO SE GENERAN SUGERIDOS "), vbInformation, "INFORMACION"
   frmobserva.Enabled = True
   frmobserva.Visible = True
   MODIFICADO = True
   dbgrdPedpro.AllowUpdate = True
Else
   MsgBox "Debe Proporcionar, La Contraseña Correcta...", vbInformation, "CONTRASEÑA"
End If
End Sub

Private Sub Form_Load()
On Error GoTo Error:
AdoPedProve.ConnectionString = cCadConex
AdoPedProve.CommandType = adCmdText
MODIFICADO = False
If Not (frmpedBod.AdoPedidos.Recordset.BOF = True And frmpedBod.AdoPedidos.Recordset.EOF) Then
  AdoPedProve.RecordSource = "SELECT * FROM Pedprove WHERE pp_pedido = '" & Trim(frmpedBod.dbgrdPed.Columns(1).Text) & "'"
Else
  AdoPedProve.RecordSource = "SELECT * FROM Pedprove"
End If
AdoPedProve.Refresh
If nOp = 1 Then  'Nuevo
     AdoPedProve.Recordset.AddNew
     AdoProv.ConnectionString = cCadConex
     AdoProv.CommandType = adCmdText
     AdoProv.RecordSource = "SELECT * FROM Catprov"
     AdoProv.Refresh
     chkCampos(0).Value = 0: chkCampos(1).Value = 0
     cmbProv.Clear
     AdoProv.Recordset.MoveFirst
     Do While Not AdoProv.Recordset.EOF
        If Not IsNull(AdoProv.Recordset!NOMPROVE) Then cmbProv.AddItem AdoProv.Recordset!NOMPROVE
        AdoProv.Recordset.MoveNext
     Loop
     lbletiquetas(1).Visible = True
     txtcampos(1).Visible = True
     cmbProv.Visible = True
     Me.dbgrdPedpro.Visible = False
ElseIf cModo = "RECIBIR" Then
    lMove = True 'Bandera que no hace el scroll al grabar si no se cicla
End If
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmpedBod.Show
End Sub

Private Sub mact_Click()
CmdRefresh_Click
End Sub

Private Sub magre_Click()
cmdAgregar_Click
End Sub

Private Sub mbarr_Click()
cmdCodBarra_Click
End Sub

Private Sub mexi_Click()
Command2_Click
End Sub

Private Sub mexp_Click()
CmdExporta_Click
End Sub

Private Sub mgrab_Click()
If cmdGrabar.Enabled = True Then
    cmdGrabar_Click
Else
    RESP = MsgBox("El Pedido ya fue Grabado, aun asi desea incrementar el inventario de Nuevo ", vbYesNo)
    If RESP = vbYes Then cmdGrabar_Click
End If
End Sub

Private Sub minve_Click()
cmdInven_Click
End Sub

Private Sub mnot_Click()
cmdReporte_Click
End Sub

Private Sub mobse_Click()
cmdobrec_Click
End Sub

Private Sub mpedsug_Click()
cmdRpteMix_Click
End Sub

Private Sub mpro_Click()
cmdsugerido_Click
End Sub

Private Sub msal_Click()
cmdRegresar_Click
End Sub

Private Sub txtCampos_GotFocus(Index As Integer)
Select Case Index
Case 1
     frmpedBod.Hide
End Select
End Sub

Private Sub txtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
     Case 27
         Unload Me
     Case 13
         KeyAscii = 0
         SendKeys "{Tab}"
End Select
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
Dim rstTie As ADODB.Recordset
Dim rsttemp As ADODB.Recordset
Dim cCadena As String
On Error GoTo Error:
Select Case Index
Case 1 'Clave del proveedor
     txtcampos(Index).Text = Trim(UCase(txtcampos(Index).Text))
     txtcampos(Index).Refresh

     'NUEVO = Ver si existen pedidos para generar un pedido por proveedor
     lpprov = True
     If nOp = 1 Then
        If txtcampos(Index).Text = "" Or IsNull(txtcampos(Index).Text) Then
            cmbProv.SetFocus
            Exit Sub
        Else
            AdoProv.Recordset.MoveFirst
            AdoProv.Recordset.Find "Prove= '" & Trim(txtcampos(Index).Text) & "'"
            If AdoProv.Recordset.EOF = True Then
                'MsgBox "No existe la clave del proveedor especificado", vbExclamation
                cmbProv.SetFocus
                Exit Sub
            End If
        End If
        chkCampos(0).Visible = True
        cmbProv.Text = AdoProv.Recordset!NOMPROVE
        'Obtengo el folio del pedido en base al proveedor Especif.
        Set rsttemp = New ADODB.Recordset
        rsttemp.ActiveConnection = cCadConex
        rsttemp.CursorType = adOpenKeyset
        rsttemp.Source = "SELECT MAX (CAST(SUBSTRING(PP_PEDIDO,5,10) AS INT)) As FolMay FROM [Pedprove] WHERE PP_proveedor = '" & txtcampos(1).Text & "'"
        rsttemp.Open
        If IsNull(rsttemp!FolMay) Then
            txtcampos(0).Text = Mid(Trim(Mid(txtcampos(1).Text, 1, 3)), 1, 3) + "-1"
        Else
            txtcampos(0).Text = txtcampos(1).Text + "-" + Trim(Str(rsttemp!FolMay + 1))
        End If
        cCond = "p_situacion = 0 AND p_proveedor = '" & txtcampos(1).Text & "'"
        'Muestro datos por default
        chkCampos(0).Value = 0
        txtcampos(3).Text = date + Time
        cmbPerCon.Visible = True
        txtcampos(2).Text = Mid(cCveDesUsu, 1, 3)
        cmbPerCon.Text = Trim(Mid(cCveDesUsu, 3))
        cmbPerCon.Enabled = False
        For N = 0 To 3
            lbletiquetas(N).Visible = True
            txtcampos(N).Visible = True
            txtcampos(N).Enabled = False
        Next
        cmbProv.Enabled = False
        dbgrdPedpro.Visible = True

        dbgrdPedsol.Visible = True
        dbgrdRec.Visible = False
        'save
        cmdGrabar.Visible = True
        'cmdsugerido.Enabled = True
     Else 'Si es consulta del pedido confirmado O recepcion
        cCond = "p_pedproveedor = '" & frmpedBod.dbgrdPed.Columns(0).Text & "'"
        For N = 0 To 4
           If N < 4 Then lbletiquetas(N).Visible = True
           txtcampos(N).Visible = True
           txtcampos(N).Enabled = False
        Next
        cmbProv.Visible = False
        chkCampos(0).Visible = True
        chkCampos(1).Visible = True
        chkCampos(0).Enabled = False
        chkCampos(1).Enabled = cModo = "RECIBIR"
        dbgrdPedpro.Visible = True
        
        'dbgrdPedpro.Width = ScaleWidth - 400
        dbgrdPedsol.Visible = True
        If cModo = "RECIBIR" Then
           'dbgrdPedpro.Width = 9735
           cmdReporte.Visible = AdoPedProve.Recordset!pp_recibe
        Else
           cmdReporte.Visible = AdoPedProve.Recordset!pp_confirma
        End If
        AdoFacturas.ConnectionString = cCadConex
        AdoFacturas.CommandType = adCmdText
        AdoFacturas.RecordSource = "SELECT * FROM [NOTAENTRADA] WHERE [pedido] = '" & txtcampos(0).Text & "'"
        AdoFacturas.Refresh
        If cModo = "RECIBIR" Then
            'dbgrdPedpro.AllowUpdate = False
            If Not AdoPedProve.Recordset!pp_recibe Then
               'No actualizo para que no se borre lo capturado por codigo de barras
               'cn.Execute "UPDATE DetalleGlobal SET dg_cantReal = dg_cantsol"
               'cn.Execute "UPDATE DetalleGlobal SET dg_cantRealp = dg_cantsolP"
               If AdoFacturas.Recordset.BOF And AdoFacturas.Recordset.EOF = True Then
                  AdoFacturas.Recordset.AddNew 'Para que no se borren los datos al ecribir en los campos del control adofacturas
                  AdoFacturas.Recordset!Pedido = txtcampos(0).Text
                  AdoFacturas.Recordset!Clavenota = "N" & Trim(txtcampos(0).Text)
                  AdoFacturas.Recordset.Update
                  AdoFacturas.Refresh
               End If
               cmdGrabar.Visible = True
               
            Else
               'For n = 0 To 3
                  txtRecib(3).Visible = True
                  lblRec(3).Visible = True
                  Me.dbgrdfactu.AllowUpdate = False
                '  txtRecib(n).Enabled = False
               'Next
               chkCampos(1).Enabled = False
               txtRecib(0).Text = AdoPedProve.Recordset!pp_fecrecibe
               'dbgrdRec.AllowUpdate = False
               cmdGrabar.Visible = True
               cmdGrabar.Enabled = False
               'cmdsugerido.Enabled = False
               cmdCodBarra.Enabled = False
               CmdAgregar.Enabled = False
               dbgrdPedpro.AllowUpdate = False
            End If
            chkCampos(1).Visible = True
        'Si es opcion ver confirmado y ya es recibido
        ElseIf AdoPedProve.Recordset!pp_recibe Then
            For N = 0 To 3
                txtRecib(N).Visible = True
                txtRecib(N).Enabled = False
            Next
            cmdGrabar.Visible = True
            cmdGrabar.Enabled = False
            'cmdsugerido.Enabled = False
            cmdCodBarra.Enabled = False
            CmdAgregar.Enabled = False
        'Si es opcion ver confirmado y no ha sido recibido
        Else
            cmdGrabar.Visible = True
            cmdGrabar.Enabled = False
            'cmdsugerido.Enabled = False
            cmdCodBarra.Enabled = False
            CmdAgregar.Enabled = False
        End If
     End If
     If cModo <> "RECIBIR" Then
        'Para crear ref. cruzada Sql Server recorro todas las tiendas
        cmdReporte.Caption = "Ped. Conf."
        Set rstTie = New ADODB.Recordset
        rstTie.Source = "SELECT * FROM cattienda"
        rstTie.ActiveConnection = cCadConex
        rstTie.Open
        'Genera la cadena del origen de datos es una referencia cruzada y
        'utiliza una vista de Sql. (DetPedTie)
        cCadena = "SELECT df_prod AS CLAVE, descripc As DESCRIPCION, str(paquetes) + ' X ' + LTRIM( str(contenid)) + ' ' + MEDIDA as ESPECIF," + Chr(13) _
        & " SUM(DF_CANTSOL) As TOTSOL,"
        While Not rstTie.EOF
            cCadena = cCadena + Chr(13) + " SUM(CASE p_sucursal WHEN '" & Trim(rstTie!ticlave) & "' THEN df_cantsol ELSE 0 END) AS " & Mid(rstTie!tidescrip, 1, 5) & ","
            rstTie.MoveNext
        Wend
        cCadena = Mid(cCadena, 1, Len(cCadena) - 1) & Chr(13) _
        & " From PedDetTie WHERE " & cCond _
        & " GROUP BY df_prod, descripc, str(paquetes) + ' X ' + LTRIM( str(contenid)) + ' ' + MEDIDA ORDER BY Descripcion"
        'Obtengo cantidades solicitadas por tienda de un proveedor especificado
        AdoPedpro.ConnectionString = cCadConex
        AdoPedpro.CommandType = adCmdText
        AdoPedpro.RecordSource = cCadena
        AdoPedpro.Refresh
        'Obtengo los folios de pedidos que solicitaron las tiendas
        AdoPedSol.ConnectionString = cCadConex
        AdoPedSol.CommandType = adCmdText
        AdoPedSol.RecordSource = "SELECT DISTINCT p_pedido as FOLIO, tidescrip AS SUCURSAL, P_fecPed As FECHA_SOL FROM PedDetTie WHERE " & cCond
        AdoPedSol.Refresh
        If AdoPedSol.Recordset.BOF And AdoPedSol.Recordset.EOF Then
           MsgBox "NO EXISTEN PEDIDOS PARA CONFIRMAR DEL PROVEEDOR ESPECIFICADO", vbExclamation
           Unload Me
           Exit Sub
        End If
     Else 'Si es recibir
        AdoPedpro.ConnectionString = cCadConex
        AdoPedpro.CommandType = adCmdText
        AdoPedpro.RecordSource = "SELECT activo , dg_producto, descripc, LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida AS MEDIDA, dg_cantsol, dg_cantsolp, dg_promocion, DETALLEGLOBAL.dg_cantreal, dg_cantrealP, dg_promocionR, dg_costo, dg_existencia FROM DetalleGlobal,tfproduc WHERE dg_pedido = '" & txtcampos(0).Text & "' AND DETALLEGLOBAL.dg_producto = TFPRODUC.consec ORDER BY DESCRIPC,MEDIDA"
        AdoPedpro.Refresh
        Me.dbgrdPedpro.Refresh
        Me.dbgrdPedsol.Visible = False
     End If
     'If cModo = "RECIBIR" Then dbgrdPedpro.Columns(3).Width = 2150
     cmdregresar.Visible = True
     cmdRpteMix.Visible = True
     'dbgrdPedpro.Columns(1).Width = 5560
     txtcampos(5).Locked = True
     PonPie
End Select
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdConAceptar_Click
End If
End Sub

Private Sub txtmodi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call MODI_PEDIDO
   FRMMODI.Enabled = False
   FRMMODI.Visible = False
ElseIf keyascci = 27 Then
   FRMMODI.Enabled = False
   FRMMODI.Visible = False
End If

End Sub

Private Sub txtobserva_KeyPress(KeyAscii As Integer)
If KeyAscii = 59 Then
   txtobserva.Text = txtobserva.Text + Chr(13)
   KeyAscii = 0
   SendKeys "{end}"
End If
End Sub

Private Sub txtRecib_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub

'Actualiza los pedidos sugeridos que formaron el pedido por proveedor
'y se generan inmediatamente un envio para que lo facturen
Private Sub ActualizaSug(PEDPROVE As String)
On Error GoTo Error:
Dim nPor
Dim rs As ADODB.Recordset
Dim RSTEMP As ADODB.Recordset
Dim rstRepart As ADODB.Recordset
Dim lSurExi As Boolean
Set rs = New ADODB.Recordset
Set RSTEMP = New ADODB.Recordset
'antes que nada se valida que no se hayan generado sus respectivos traslados
'Se recorre todo el detalle del pedido global y se van poniendo cantidades a los sugeridos en base al
'porcentaje que representa lo solicitado en el sug. a lo recibido en el ped. por proveedor
cn.BeginTrans
If Trim(txtcampos(1).Text) <> "JAR" Then
  AdoPedpro.Recordset.MoveFirst
  'SE DETERMINA SI SE VA A SURTIR TAMBIEN DE BODEGA
  'rs.Open "SELECT SUM(dg_cantsol) AS SurBod FROM detalleglobal WHERE dg_pedido = '" & txtCampos(0).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
  'If rs!SURBOD > 0 Then
  '   lSurExi = False
  'Else
  '   lSurExi = True
  'End If
  'lSurExi = (rs!SURBOD = 0)
  lSurExi = True
  rs.Close
  'SE DEBE SURTIR CON BASE A LAS EXISTENCIAS, YA SE SUMO LO QUE SE RECIBIO
  While Not AdoPedpro.Recordset.EOF
    If lSurExi Then  'Todo se va a surtir de las existencias de bodega
       rs.Open "SELECT generado, df_pedido,df_prod, df_cantsol, df_cantreal FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PEDPROVE & "' AND DF_PROD = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND p_sucursal <> 10 AND df_sugerido = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
       nCanSug = 0
       nSolOfi = 0
       'SI existe un sugerido de oficinas centrales se reparte entre las demas tiendas
       Set rstRepart = New ADODB.Recordset
       rstRepart.Open "SELECT * FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PEDPROVE & "' AND DF_PROD = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND P_SUCURSAL = 10", cn, adOpenKeyset, adLockOptimistic, adCmdText
       If rstRepart.RecordCount > 0 Then
          nSolOfi = Round(rstRepart!df_cantsol / rs.RecordCount, 0)
       End If
       While Not rs.EOF
           RSTEMP.Open "SELECT SUM(df_cantsol) AS Cansol, InCant FROM pedidos,detallefactura,inventario WHERE p_pedproveedor = '" & txtcampos(0).Text & "' AND DF_PEDIDO = P_PEDIDO AND DF_PROD = INPROD AND df_prod = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND df_sugerido = 1 GROUP BY InCant", cn, adOpenKeyset, adLockOptimistic, adCmdText
           nPor = 100
           If RSTEMP!InCant < RSTEMP!cansol And RSTEMP!InCant > 0 Then 'Se surte equitativamente entre todas las tiendas lo que hay en el inventario
              nPor = (rs!df_cantsol) * 100 / RSTEMP!InCant
              ncan = Round(RSTEMP!cansol * nPor / 100, 0)
           Else
              ncan = Round((rs!df_cantsol + nSolOfi) * nPor / 100, 0)
           End If
           nCanSug = nCanSug + ncan
           cProAnt = rs!df_prod
           If nCanSug > RSTEMP!InCant Then  'Cuando ya no alcanza el inventario para surtir
              'If Not rs!generado Then
                 rs!df_cantreal = RSTEMP!InCant - (nCanSug - ncan)
                 'rs!generado = 1
                 rs.Update
                 rs.MoveLast 'se agotaron las existencias se mueve al final paraque ya no siga asignando
             ' End If
           ElseIf nCanSug > RSTEMP!cansol Then  'Cuando hay mas inventario de lo solicitado
              'If Not rs!generado Then
                 rs!df_cantreal = RSTEMP!cansol - (nCanSug - ncan)
                 'rs!generado = 1
                 rs.Update
              'End If
           Else
              'If Not rs!generado Then
                 rs!df_cantreal = ncan
                 'rs!generado = 1
                 rs.Update
              'End If
           End If
           rs.MoveNext
           RSTEMP.Close
       Wend
       rs.Close
    Else  'Se reparte de lo recibido en el pedprove
        If AdoPedpro.Recordset!dg_cantreal > 0 Then
           'Aunque se repite el codigo es mas rapido porque no procesa los que la cantidad recibida es mayor a cero
           'obtengo todas las tiendas que pidieron el producto
           rs.Open "SELECT generado,df_pedido,df_prod, df_cantsol, df_cantreal FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PEDPROVE & "' AND DF_PROD = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND p_sucursal <> 10 AND df_sugerido = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
           nCanSug = 0
           nSolOfi = 0
           Set rstRepart = New ADODB.Recordset
           'Obtengo el total de la cantidad solicitada de los sugeridos ya que no coincide con lo solicitado del pedprove
           rstRepart.Open "SELECT sum(df_cantsol) as SolSug FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PEDPROVE & "' AND DF_PROD = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND df_sugerido = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
           nCanSolSug = IIf(IsNull(rstRepart), 0, rstRepart!SolSug)
           rstRepart.Close
           
           'SI existe un sugerido de oficinas centrales se reparte entre las demas tiendas
           rstRepart.Open "SELECT * FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & PEDPROVE & "' AND DF_PROD = '" & AdoPedpro.Recordset!DG_PRODUCTO & "' AND P_SUCURSAL = 10", cn, adOpenKeyset, adLockOptimistic, adCmdText
           If rstRepart.RecordCount > 0 Then
              If rs.RecordCount > 0 Then nSolOfi = Round(rstRepart!df_cantsol / rs.RecordCount, 0)
           End If
           cProAnt = ""
           While Not rs.EOF
                StbMensajes.SimpleText = Space(25) & "Prorrateando producto del pedido sugerido " & rs!df_pedido
                StbMensajes.Refresh

                If cProAnt <> rs!df_prod Then
                    nCanSug = 0
                End If
                nCanRec = AdoPedpro.Recordset!dg_cantreal + AdoPedpro.Recordset!DG_PROMOCIONR
                nPor = (rs!df_cantsol + nSolOfi) * 100 / IIf(nCanSolSug = 0, 1, nCanSolSug)
                ncan = Round(nCanRec * IIf(nPor = 0, 1, nPor) / 100, 0)
                nCanSug = nCanSug + ncan
                cProAnt = rs!df_prod
                If nCanSug > nCanRec Then
                    'If Not rs!generado Then
                       rs!df_cantreal = nCanRec - (nCanSug - ncan)
                       'rs!generado = 1
                    'End If
                Else
                    'If Not rs!generado Then
                      ' rs!generado = 1
                       rs!df_cantreal = ncan
                    'End If
                End If
                rs.Update
                rs.MoveNext
            Wend
            'En el caso que aun sobre producto por lo regular es uno
            If nCanSug < nCanRec And rs.RecordCount > 0 Then
               rs.MoveLast
               If Not rs!generado Then
                  rs!df_cantreal = rs!df_cantreal + (nCanRec - nCanSug)
                  'rs!generado = 1
                  rs.Update
               End If
            End If
            rs.Close
        End If
    End If
    AdoPedpro.Recordset.MoveNext
  Wend
End If
'cn.CommitTrans
'Genero envios con los sugeridos y los marco como recibidos
MsgBox "Este Proceso Requiere de Uso Exclusivo del Archivo de Traslados, presione Enter para Continuar...", vbInformation
'cn.BeginTrans
RSTEMP.Open "select grabando from cattienda where ticlave = 3", cn, adOpenKeyset, adLockOptimistic, adCmdText
If RSTEMP!grabando Then
   MsgBox "En este momento se esta Generando recibo de otro Proveedor,Salga del Proceso y Vuelva a intentarlo", vbInformation
   RSTEMP.Close
   Exit Sub
Else
    RSTEMP!grabando = 1
    RSTEMP.Update
    RSTEMP.Close
End If
'cn.BeginTrans
rs.Open "SELECT * FROM pedidos WHERE p_pedproveedor = '" & PEDPROVE & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
While Not rs.EOF
'  If Trim(rs!p_sucursal) <> "14" And Trim(rs!p_sucursal) <> "15" And Trim(rs!p_sucursal) <> "12" And Trim(rs!p_sucursal) <> "5" And Trim(rs!p_sucursal) <> "4" And Trim(rs!p_sucursal) <> "10" Then
  'No se generan traslados para Hidalgo y Oficinas centrales
  If Trim(rs!p_sucursal) <> "4" And Trim(rs!p_sucursal) <> "10" Then
     StbMensajes.SimpleText = Space(25) & "Espere un momento, Generando envio del pedido sugerido " & rs!p_Pedido
     StbMensajes.Refresh
     RSTEMP.Open "SELECT MAX (CAST(SUBSTRING(t_clave,4,10) AS INT)) As FolTra FROM Traslados WHERE SUBSTRING(t_clave,1,3) = 'S" & Trim(Mid(cSucursal, 3, 5)) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
     cfoltra = IIf(IsNull(RSTEMP!FolTra), "S" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + "1", "S" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + Trim(Str(RSTEMP!FolTra + 1)))
     RSTEMP.Close
     nfoltie = 0
     'Agrego el traslado
     If Not rs!P_recibido Then
        cn.Execute "INSERT INTO traslados(t_clave, t_fecha, t_tipo, t_sucursalemisor, t_sucursalreceptor, t_perfle, t_foliotie ) VALUES ('" & cfoltra & "','" & date + Time & "',0,'" & Trim(Mid(cSucursal, 1, 3)) & "','" & rs!p_sucursal & "','3'," & nfoltie & ")"
       If Trim(txtcampos(1).Text = "JAR") Then
           cn.Execute "INSERT INTO DetalleTraslado (dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido) SELECT  dt_clave = '" & _
                      cfoltra & "',df_prod,df_cantreal,df_cantrealp, pedido = '" & rs!p_Pedido & "' FROM detallefactura WHERE df_pedido = '" & rs!p_Pedido & "'AND df_sugerido = 1"
       Else  'En Jarcieria quien toma el control total para el surtimiento es el facturista asi es que se le pone lo solicitado
           cn.Execute "INSERT INTO DetalleTraslado (dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido) SELECT  dt_clave = '" & _
                      cfoltra & "',df_prod,df_cantreal,df_cantrealp, pedido = '" & rs!p_Pedido & "' FROM detallefactura WHERE df_pedido = '" & rs!p_Pedido & "' AND df_sugerido = 1"
       End If
     End If
     'Cargo precios del traslado; Precio a franquicias (PRECIO4  de PREPROD)
     'If Val(rs!p_sucursal) = 5 Or Val(rs!p_sucursal) = 12 Or Val(rs!p_sucursal) = 13 Or Val(rs!p_sucursal) = 15 Or Val(rs!p_sucursal) = 14 Or Val(rs!p_sucursal) = 27 Then
     'MsgBox rstemp!franquicia
     'MsgBox rstemp!tidescrip
     RSTEMP.Open "SELECT tidescrip , franquicia FROM Cattienda WHERE ticlave = '" & Trim(rs!p_sucursal) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
     If Not rs!P_recibido Then
        'AQUI SE PONEN LOS COSTOS Y PRECIOS DE LOS TRASLADOS
         'PARA FRANQUICIAS COSTO =  COSTO, PRECIOVENTA = PRECIO4
         'PARA TIENDAS COSTO = COSTO, PRECIOVENTA = PRECIO2
         'DE ESA MANERA SE DEBEN JALAR EN FORMA DE TRASLADOS
        If RSTEMP!franquicia Then
            'cn.Execute "UPDATE DETALLETRASLADO SET DETALLETRASLADO.dt_costo = TFPRODUC.precosto, DETALLETRASLADO.dt_costoP = TFPRODUC.precosto / TFPRODUC.paquetes, DETALLETRASLADO.dt_venta = PREPROD.precio4, DETALLETRASLADO.dt_ventaP = PREPROD.precio4 / TFPRODUC.paquetes FROM PREPROD,TFPRODUC WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & cfoltra & "' AND tfproduc.ACTIVO = 1"
            CAD = "UPDATE DETALLETRASLADO SET DETALLETRASLADO.dt_costo = TFPRODUC.precosto, DETALLETRASLADO.dt_costoP = TFPRODUC.precosto / TFPRODUC.paquetes, DETALLETRASLADO.dt_venta = PREPROD.precio4, DETALLETRASLADO.dt_ventaP = PREPROD.precio4 / TFPRODUC.paquetes FROM PREPROD,TFPRODUC WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto  AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & cfoltra & "' "
            'MsgBox cad
            cn.Execute CAD
        Else 'Precio a tiendas (Precosto de TFPRODUC)
            'cn.Execute "UPDATE DETALLETRASLADO SET dt_costo = TFPRODUC.Precosto, dt_costoP = TFPRODUC.Precosto / TFPRODUC.paquetes , dt_venta = PREPROD.precio2, dt_ventaP = PREPROD.precio1 FROM TFPRODUC,PREPROD WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & cfoltra & "' AND tfproduc.ACTIVO = 1"
            CAD = "UPDATE DETALLETRASLADO SET dt_costo = TFPRODUC.Precosto, dt_costoP = TFPRODUC.Precosto / TFPRODUC.paquetes , dt_venta = PREPROD.precio2, dt_ventaP = PREPROD.precio1 FROM TFPRODUC,PREPROD WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & cfoltra & "' "
            cn.Execute CAD
        End If
        rs!p_traslado = cfoltra
        MsgBox "SE GENERO EL TRASLADO PARA " & Trim(RSTEMP!tidescrip) & " CON LA CLAVE  " & CStr(cfoltra) & "  ", vbInformation
        'rstemp.Close
    End If
    RSTEMP.Close
  End If
    If Not rs!P_recibido Then
        rs!P_recibido = 1
        rs!p_fecentreal = txtRecib(0).Text
        rs.Update
        rs.MoveNext
    End If
Wend
AdoPedProve.Recordset!pp_sugeridos = True
AdoPedProve.Recordset.Update
'cn.Execute "update pedprove set pp_sugeridos = 1 where pp_pedido = '" & Trim(Me.AdoPedProve.Recordset!pp_pedido) & "'"
RSTEMP.Open "select grabando from cattienda where ticlave = 3"
RSTEMP!grabando = 0
RSTEMP.Update
RSTEMP.Close
cn.CommitTrans
Exit Sub
Error:
  MsgBox " Se genero un error al generar de Sugeridos a Traslados, VUELVA A REALIZAR EL PROCESO", vbCritical
RSTEMP.Open "select grabando from cattienda where ticlave = 3"
RSTEMP!grabando = 0
RSTEMP.Update
RSTEMP.Close
cn.RollbackTrans
End Sub


Private Sub ActSugnew(PEDPROVE As String)
'On Error GoTo error:
If Trim(txtcampos(1).Text) <> "JAR" Then
  AdoPedpro.Recordset.MoveFirst
  While Not AdoPedpro.Recordset.EOF
      Call PRORRATEA1(PEDPROVE)
      AdoPedpro.Recordset.MoveNext
  Wend
  'Y POR ULTIMO SE ACTUALIZAN LAS FECHAS A LOS PEDIDOS SUGERIDOS DE LAS TIENDAS
  CAD = "UPDATE PEDIDOS SET P_FECENTREAL = '" & date + Time & "' WHERE p_pedproveedor = '" & Trim(PEDPROVE) & "'"
  Me.StbMensajes.SimpleText = "Espere actualizando Fechas de Recibo en Pedidos Sugeridos..."
  cn.Execute CAD
  Call GENERAENVIOS(PEDPROVE)
End If
Exit Sub
Error:
   cn.RollbackTrans
End Sub


Private Sub PRORRATEA(PEDPROVE As String)
Dim CAJASINV As Integer
Dim CAJASPEDIDAS As Integer
Dim rspro As ADODB.Recordset
Set rspro = New ADODB.Recordset
Dim rsinv As ADODB.Recordset
Set rsinv = New ADODB.Recordset
CAD = "SELECT df_pedido,df_prod, df_cantsol, df_cantreal,df_cantrealp FROM pedidos,detallefactura WHERE p_pedido = df_pedido  AND P_PEDPROVEEDOR = '" & Trim(PEDPROVE) & "' AND DF_PROD = '" & Trim(AdoPedpro.Recordset!DG_PRODUCTO) & "' AND p_sucursal <> 10  "
rspro.Open CAD, cn, adOpenKeyset, adLockOptimistic, adCmdText
If rspro.EOF Then
   Exit Sub
End If
CAD = "SELECT INCANT, INCANTPZA FROM INVENTARIO WHERE INPROD = '" & Trim(AdoPedpro.Recordset!DG_PRODUCTO) & "'  order by df_cantsol "
rsinv.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
If Not rsinv.EOF Then
  CAJASINV = rsinv!InCant
End If
Set rsinv = Nothing
CAJASPEDIDAS = 0
rspro.MoveFirst
While Not rspro.EOF
   CAJASPEDIDAS = CAJASPEDIDAS + rspro!df_cantsol
   rspro.MoveNext
Wend
If CAJASPEDIDAS < CAJASINV Then
    NOPRORRATEA = True
End If
INVENTARIOT = CAJASINV
rspro.MoveFirst
LLEVO = 0
While Not rspro.EOF
  t = rspro!df_pedido
   'PARA QUE NO HAYA BRONCAS SI NO ES NECESARIO PRORRATEAR
   If NOPRORRATEA Then
      cantpro = rspro!df_cantsol
   Else
        If rspro!df_cantsol > 0 Then
            cantemp = (rspro!df_cantsol / CAJASPEDIDAS) * INVENTARIOT
            If cantemp < 0.99 Then
                cantpro = 1
            Else
                cantpro = Round(cantemp)
            End If
        Else
            cantpro = 0
        End If
   End If
   'un momento si es menor lo que solicita a lo que sale
   If rspro!df_cantsol < cantpro Then
       rspro!df_cantreal = rspro!df_cantsol
       cantidad = rspro!df_cantsol
       LLEVO = LLEVO + rspro!df_cantsol
   Else
       rspro!df_cantreal = cantpro
       LLEVO = LLEVO + cantpro
       cantidad = cantpro
   End If
   CAJASINV = CAJASINV - cantidad
   If CAJASINV < 0 Then
      't = CAJASINV + cantidad
      'tt =
      rspro!df_cantreal = 0
   End If
   'MsgBox rspro!df_pedido & " Pedido : " & rspro!df_cantsol & "  dado : " & rspro!df_cantreal
   rspro!df_cantrealP = 0
   rspro.Update
   rspro.MoveNext
Wend
rspro.Close
'SI QUEDO UN POCO DE INVENTARIO
'If CAJASINV > 0 Then
  ' If LLEVO < INVENTARIOT Then
      'ES NECESARIO VOLVER A REASIGNAR PARA PODER A COMPLETAR EL PEDIDO
      'SE SELECCIONAN TODOS LOS PEDIDOS Y SE LE EMPIEZA A AGREGAR A UNO POR UNO
      'HASTA QUE SE ACABE EL INVENTARIO
   '     cad = "SELECT df_pedido,df_prod, df_cantsol, df_cantreal,df_cantrealp FROM detallefactura,pedidos WHERE df_pedido = p_pedido AND P_PEDPROVEEDOR = '" & Trim(PEDPROVE) & "' AND DF_PROD = '" & Trim(AdoPedpro.Recordset!DG_PRODUCTO) & "' AND p_sucursal <> 10  ORDER BY DF_CANTSOL "
     '   rspro.Open cad, cn, adOpenKeyset, adLockOptimistic, adCmdText
    '    If rspro.EOF Then
     '       Exit Sub
      '  End If
      
   ' End If
'End If

End Sub
           
           
Private Sub PRORRATEA1(PEDPROVE As String)
Dim CAJASINV As Integer
Dim CAJASPEDIDAS As Integer
Dim rspro As ADODB.Recordset
Set rspro = New ADODB.Recordset
Dim rsinv As ADODB.Recordset
Set rsinv = New ADODB.Recordset
CAD = "SELECT df_pedido,df_prod, df_cantsol, df_cantreal,df_cantrealp FROM pedidos,detallefactura WHERE p_pedido = df_pedido  AND P_PEDPROVEEDOR = '" & Trim(PEDPROVE) & "' AND DF_PROD = '" & Trim(AdoPedpro.Recordset!DG_PRODUCTO) & "' AND p_sucursal <> 10  "
rspro.Open CAD, cn, adOpenKeyset, adLockOptimistic, adCmdText
If rspro.EOF Then
   Exit Sub
End If
CAD = "SELECT INCANT, INCANTPZA FROM INVENTARIO WHERE INPROD = '" & Trim(AdoPedpro.Recordset!DG_PRODUCTO) & "'  "
rsinv.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
If Not rsinv.EOF Then
  CAJASINV = rsinv!InCant
End If
Set rsinv = Nothing
CAJASPEDIDAS = 0
rspro.MoveFirst
While Not rspro.EOF
   CAJASPEDIDAS = CAJASPEDIDAS + rspro!df_cantsol
   rspro.MoveNext
Wend
If CAJASPEDIDAS < CAJASINV Then
    NOPRORRATEA = True
End If
INVENTARIOT = CAJASINV
rspro.MoveFirst
LLEVO = 0
For i = 1 To 50
  descripcs(i) = ""
  cantidads(i) = 0
  tasas(i) = 0
  costoss(i) = 0
  costosp(i) = 0
Next
i = 1
rspro.MoveFirst
LLEVO = 0
While Not rspro.EOF
  ' SE SUBE TODO A UNA ARREGLO
  descripcs(i) = rspro!df_pedido
  cantidads(i) = rspro!df_cantsol
  If CAJASPEDIDAS < 1 Then
      tasas(i) = 0
  Else
      tasas(i) = (rspro!df_cantsol / CAJASPEDIDAS) * 100
  End If
  If NOPRORRATEA Then
     costoss(i) = rspro!df_cantsol
  Else
    costoss(i) = Round(((tasas(i) * CAJASINV) / 100) + 0.1)
  End If
  rspro.MoveNext
  LLEVO = LLEVO + costoss(i)
  i = i + 1
Wend
'1 condicion , no debe de sobrar cajas
If LLEVO < CAJASINV Then
   For j = 1 To i
        If cantidads(j) > costoss(j) Then
            costoss(j) = costoss(j) + 1
            LLEVO = LLEVO - 1
        End If
        If LLEVO < 1 Then
           Exit For
        End If
   Next
End If
'2 no debe de haber de mas
If LLEVO > CAJASINV Then
    difer = LLEVO - CAJASINV
    For j = 1 To i
        'MsgBox J
        'MsgBox costoss(J)
        If costoss(j) > 1 Then
           costoss(j) = costoss(j) - 1
           difer = difer - 1
        End If
        If difer < 1 Then
           Exit For
        End If
    Next
End If
enviadas = 0
For j = 1 To i
  enviadas = enviadas + costoss(j)
Next
'3 se comprueba lo que se pidio y lo que se queda en inventario
If CAJASINV >= CAJASPEDIDAS Then
   'hubo sifuciente inventario
   If enviadas < CAJASPEDIDAS Then
      For Y = 1 To i
          costoss(Y) = costoss(Y) + 1
          enviadas = enviadas + 1
          If enviadas = CAJASPEDIDAS Then
             Exit For
          End If
      Next
   End If
Else ' se pidio justo o realizo el prorrateo
   If enviadas < CAJASINV Then
       For Y = 1 To i
          costoss(Y) = costoss(Y) + 1
          enviadas = enviadas + 1
          If enviadas = CAJASINV Then
             Exit For
          End If
      Next
   End If
End If
'AHORA SE PASA A LA TABLA
rspro.MoveFirst
j = 1
While Not rspro.EOF
    rspro!df_cantreal = costoss(j)
    rspro.Update
    rspro.MoveNext
    j = j + 1
Wend
Set rspro = Nothing

End Sub
           
           
Private Sub GENERAENVIOS(PEDPROVE As String)
'On Error GoTo error:
Dim nPor
Dim total As Integer
Dim i As Integer
Dim Pedido As String
Dim fecha As Date
Dim rs As ADODB.Recordset
Dim RSTEMP As ADODB.Recordset
Dim tienda As Integer
Dim franquicia As Boolean
Set rs = New ADODB.Recordset
Me.StbMensajes.SimpleText = "Espere un momento... Detectando Tipo de Proveedor"
StbMensajes.Refresh
PROVEEDOR = Mid(PEDPROVE, 1, 3)
rs.Open "SELECT * FROM catprov WHERE prove = '" & Trim(PROVEEDOR) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
If rs.EOF Then
   MsgBox "Imposible Detectar Tipo de Proveedor, Se Tomara la Opcion por Default", vbInformation
   tipo = 2 ' NORMAL
Else
   If rs!Volumen Then
      tipo = 1 ' POR VOLUMEN
   Else
      tipo = 2 ' NORMAL
   End If
End If
frecuencia = rs!frecuencia
rs.Close
rs.Open "SELECT * FROM pedidos WHERE p_pedproveedor = '" & Trim(PEDPROVE) & "' order by p_pedido ", cn, adOpenKeyset, adLockOptimistic, adCmdText
Set RSTEMP = New ADODB.Recordset
While Not rs.EOF
  If Trim(rs!p_sucursal) <> "4" And Trim(rs!p_sucursal) <> "10" Then
     StbMensajes.SimpleText = Space(25) & "Espere un momento, Generando envio del pedido sugerido " & rs!p_Pedido
     StbMensajes.Refresh
     RSTEMP.Open "SELECT  tidescrip , franquicia FROM Cattienda WHERE ticlave = '" & Trim(rs!p_sucursal) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
     If RSTEMP.EOF Then
        franquicia = False 'POR DEFAULT
     Else
        franquicia = RSTEMP!franquicia
     End If
     tienda = rs!p_sucursal
     Pedido = rs!p_Pedido
     NOMBRE = RSTEMP!tidescrip
     RSTEMP.Close
     'MOY
     If Not rs!P_recibido Then
        If tipo = 2 Then ' PEDIDO NORMAL
                fecha = date
                'pedido, fecha, tipo de pedido, numero de traslado, total de traslados
                Pedido = rs!p_Pedido
                cfoltra = agregatraslado(Pedido, fecha, 2, 1, 1, franquicia, tienda)
                rs!p_traslado = cfoltra
               ' MsgBox "SE GENERO EL TRASLADO PARA " & Trim(NOMBRE) & " CON LA CLAVE  " & CStr(cfoltra) & "  ", vbInformation
        Else ' POR VOLUMEN
                total = IIf(frecuencia = 15, 6, 3)
                fechaultima = date + (total * 2)
                For i = 1 To total
                    fecha = fecha + 3 'Round(frecuencia / total)
                    fecha = fechaultima - (i * 2)
                    cfoltra = agregatraslado(Pedido, fecha, 1, i, total, franquicia, tienda)
                    rs!p_traslado = cfoltra
                    rs.Update
                Next
        End If
     End If
  End If 'FIN DE GENERAR EL ENVIO
  If Not rs!P_recibido Then
        rs!P_recibido = 1
        rs!p_fecentreal = txtRecib(0).Text
        rs.Update
  End If
  rs.MoveNext
Wend
AdoPedProve.Recordset!pp_sugeridos = True
AdoPedProve.Recordset.Update
MsgBox "Proceso de Generacion de Envios Finalizado" & vbCrLf & "Informe al Area de Facturacion a Tiendas ", vbInformation, "SUGERIDOS"
Exit Sub
Error:
  MsgBox " Se genero un error al generar de Sugeridos a Traslados, VUELVA A REALIZAR EL PROCESO", vbCritical
MsgBox "Se genero un Error en el proceso de generacion de envios, por favor, vuelva a dar click en [sug]", vbInformation, "SUGERIDOS"
'cn.RollbackTrans
End Sub

'Pedido = Clave del pedido del cual se generará el traslado
'fecha  = fecha con la que se generará el traslado
'totalped  = que tipo de pedido es 1-> Normal  2->Volúmen
'Numpedido = Numero de traslado de la serie
'totalped = Total de traslados a generarse
'Franquicia = Determina si es tienda o franquicia
'Tienda = Es la clave de la tienda
Private Function agregatraslado(Pedido As String, fecha As Date, tipo As Integer, NUMPEDIDO As Integer, TOTALPED As Integer, franquicia As Boolean, tienda As Integer) As String
Dim RSTEMP As ADODB.Recordset
Set RSTEMP = New ADODB.Recordset
Dim cfoltra As String
Me.StbMensajes.SimpleText = "AGREGANDO NUEVO TRASLADO..."
RSTEMP.Open "SELECT MAX (CAST(SUBSTRING(t_clave,4,10) AS INT)) As FolTra FROM Traslados WHERE SUBSTRING(t_clave,1,3) = 'S" & Trim(Mid(cSucursal, 3, 5)) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
cfoltra = IIf(IsNull(RSTEMP!FolTra), "S" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + "1", "S" + Mid(Trim(Mid(cSucursal, 3)), 1, 2) + Trim(Str(RSTEMP!FolTra + 1)))
Set RSTEMP = Nothing
nfoltie = 0
'SE DETERMINA SI ES DE VOLUMEN O NO
If TOTALPED > 1 Then
    CAD = "INSERT INTO traslados(t_clave, t_fecha, t_tipo, t_sucursalemisor, t_sucursalreceptor, t_perfle, t_foliotie,T_OBSERVA ) VALUES ('" & _
                    cfoltra & "','" & fecha + Time & "',0,'" & Trim(Mid(cSucursal, 1, 3)) & "','" & tienda & "','3'," & nfoltie & ",'VOLUMEN' )"
Else
   CAD = "INSERT INTO traslados(t_clave, t_fecha, t_tipo, t_sucursalemisor, t_sucursalreceptor, t_perfle, t_foliotie,T_OBSERVA ) VALUES ('" & _
                    cfoltra & "','" & fecha + Time & "',0,'" & Trim(Mid(cSucursal, 1, 3)) & "','" & tienda & "','3'," & nfoltie & ",'' )"
End If
cn.Execute CAD
StbMensajes.SimpleText = "Agregando el detalle del traslado " & cfoltra
Call AGREGADETALLETRAS(Pedido, franquicia, cfoltra, NUMPEDIDO, TOTALPED)
agregatraslado = cfoltra
'EL COSTO
If franquicia Then
   CAD = "UPDATE DETALLETRASLADO SET DETALLETRASLADO.dt_costo = TFPRODUC.precosto, DETALLETRASLADO.dt_costoP = TFPRODUC.precosto / TFPRODUC.paquetes, DETALLETRASLADO.dt_venta = PREPROD.precio4, DETALLETRASLADO.dt_ventaP = PREPROD.precio4 / TFPRODUC.paquetes FROM PREPROD,TFPRODUC WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto  AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & cfoltra & "' "
Else
   CAD = "UPDATE DETALLETRASLADO SET dt_costo = TFPRODUC.Precosto, dt_costoP = TFPRODUC.Precosto / TFPRODUC.paquetes , dt_venta = PREPROD.precio2, dt_ventaP = PREPROD.precio1 FROM TFPRODUC,PREPROD WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & cfoltra & "' "
End If
cn.Execute CAD
End Function

Private Function AGREGADETALLETRAS(Pedido As String, franquicia As Boolean, cfoltra As String, NUMPEDIDO As Integer, TOTALPED As Integer) As String
Dim RSTEMP As ADODB.Recordset
Set RSTEMP = New ADODB.Recordset
RSTEMP.Open "SELECT *  FROM detallefactura  WHERE df_sugerido = 1 and  df_pedido =  '" & Trim(Pedido) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
While Not RSTEMP.EOF
       If RSTEMP!df_cantreal > 0 Then
        End If
       canti = RSTEMP!df_cantreal
       caNtip = RSTEMP!df_cantrealP
       cantidadgenerada = Int(canti / TOTALPED)
       CANTIDADGENERADAP = Int(caNtip / TOTALPED)
       CANTIDADGENERADAP = 0
       If NUMPEDIDO = TOTALPED Then
            'EL ULTIMO
            If cantidadgenerada < (canti / TOTALPED) Then
                cantidadgenerada = cantidadgenerada + 1
            End If
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
            'EN EL ULTIMO TRASLADO SE VA EL RESTO SE CUADRA
            rs.Open "SELECT sum(dt_cantidad) as llevo  FROM detalletraslado ,traslados WHERE t_clave = dt_clave and t_tipo = 0 and dt_pedido =  '" & Trim(Pedido) & "' and dt_producto = '" & Trim(RSTEMP!df_prod) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
            llevox = rs!LLEVO
            Set rs = Nothing
            difer = canti - (llevox + cantidadgenerada)
            If difer > 0 Then
                 cantidadgenerada = cantidadgenerada + difer
            End If
       End If
       
       CAD = "INSERT INTO DetalleTraslado (dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido) " & _
             "VALUES( '" & Trim(cfoltra) & "','" & Trim(RSTEMP!df_prod) & "'," & cantidadgenerada & "," & CANTIDADGENERADAP & ",'" & Trim(Pedido) & "')"
       cn.Execute CAD
       RSTEMP.MoveNext
Wend
Set RSTEMP = Nothing
cn.Execute CAD
End Function

Private Sub PonPie()
On Error Resume Next
Dim rsttemp As ADODB.Recordset
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT COUNT(dg_producto) AS TOTPROD, SUM(dg_cantreal+dg_promocionr) AS TOTCAJAS, SUM(dg_cantrealP) AS TOTPIEZAS  FROM Detalleglobal WHERE Dg_pedido = '" & txtcampos(0).Text & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
    lblCajas.Visible = True
    If IsNull(rsttemp!TOTCAJAS) Or IsNull(rsttemp!TOTPIEZAS) Then
       lblCajas.Caption = "TOTAL DE PROD: 0   CAJAS: 0   PIEZAS: 0"
    Else
       lblCajas.Caption = "PROD: " & CStr(rsttemp!TOTPROD) & Space(3) & "CAJAS: " & CStr(rsttemp!TOTCAJAS) & Space(3) & "PIEZAS: " & CStr(rsttemp!TOTPIEZAS)
    End If
    lblCajas.Refresh
End Sub

