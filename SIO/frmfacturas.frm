VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFacturas 
   BackColor       =   &H8000000B&
   Caption         =   "Facturas generadas"
   ClientHeight    =   8310
   ClientLeft      =   255
   ClientTop       =   1230
   ClientWidth     =   8880
   Icon            =   "frmfacturas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.Frame FraAbono 
      BackColor       =   &H80000004&
      Caption         =   "Abonos a la factura nn de la serie B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2655
      Left            =   120
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   11775
      Begin MSAdodcLib.Adodc AdoAbono 
         Height          =   330
         Left            =   2880
         Top             =   1200
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
         Caption         =   "AdoAbonos"
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
      Begin VB.CommandButton CmdAbonp 
         Caption         =   "Regresar"
         Height          =   425
         Index           =   1
         Left            =   10320
         Picture         =   "frmfacturas.frx":400A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Ocultar información de Parcialidades"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton CmdAbonp 
         Caption         =   "Actualizar"
         Height          =   425
         Index           =   0
         Left            =   10320
         Picture         =   "frmfacturas.frx":417C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Actualiza información de Parcialidades"
         Top             =   360
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dbgrdAbono 
         Bindings        =   "frmfacturas.frx":427E
         Height          =   2295
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   13817798
         BorderStyle     =   0
         ForeColor       =   8388608
         HeadLines       =   1.4
         RowHeight       =   15
         TabAction       =   2
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
            DataField       =   "SERIE"
            Caption         =   "Serie"
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
            DataField       =   "factura"
            Caption         =   "Factura"
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
            DataField       =   "totfac"
            Caption         =   "Total a Pagar"
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
         BeginProperty Column03 
            DataField       =   "totabono"
            Caption         =   "Total Abonos"
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
         BeginProperty Column04 
            DataField       =   "debe"
            Caption         =   "Por Pagar"
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
         BeginProperty Column06 
            DataField       =   "importe"
            Caption         =   "Abono"
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
            DataField       =   "tipopag"
            Caption         =   "Tipo Pago o Banco"
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
            DataField       =   "numero"
            Caption         =   "Num-Cheque"
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
         BeginProperty Column09 
            DataField       =   "posfechado"
            Caption         =   "Posfechado"
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
         BeginProperty Column10 
            DataField       =   "comenta"
            Caption         =   "                             Comentarios"
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
               DividerStyle    =   3
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               DividerStyle    =   3
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   4830.236
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fradescripcion 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11775
      Begin VB.CheckBox chkcobro 
         Caption         =   "Fecha de Cobro"
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
         Left            =   7200
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkCancelado 
         Caption         =   "Factura Cancelada"
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
         Left            =   4200
         TabIndex        =   26
         Top             =   525
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   1
         Left            =   10080
         TabIndex        =   31
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   61145091
         CurrentDate     =   36892
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   0
         Left            =   8400
         TabIndex        =   30
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   61145091
         CurrentDate     =   36892
      End
      Begin VB.Label lblSit 
         BackColor       =   &H80000004&
         Caption         =   "Total de facturas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 999,999"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label lblSituacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   195
         Width           =   5655
      End
      Begin VB.Label lblSit 
         BackColor       =   &H80000004&
         Caption         =   "Factura a :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   200
         Width           =   975
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Fecha Inicial"
         Height          =   255
         Index           =   0
         Left            =   8400
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Fecha Final"
         Height          =   255
         Index           =   1
         Left            =   10080
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox CR1 
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   46
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Frame fraCon 
      BackColor       =   &H00C0C000&
      Caption         =   "Contraseña de acceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3600
      TabIndex        =   21
      Top             =   3720
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
         BackColor       =   &H00C0C000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FraPolizas 
      Caption         =   "Pólizas de ingresos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   39
      Top             =   1680
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid dbgrdpol 
         Bindings        =   "frmfacturas.frx":4295
         Height          =   5535
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   9763
         _Version        =   393216
         AllowArrows     =   -1  'True
         BackColor       =   -2147483624
         HeadLines       =   1.3
         RowHeight       =   15
         TabAction       =   2
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
         Caption         =   "VENTAS DIARIAS DEL 01/10/02 AL 01/10/02"
         ColumnCount     =   17
         BeginProperty Column00 
            DataField       =   "folio"
            Caption         =   "Folio"
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
         BeginProperty Column02 
            DataField       =   "Bancos"
            Caption         =   "Banamex"
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
         BeginProperty Column03 
            DataField       =   "efectivo"
            Caption         =   "Eftvo.(Stand)"
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
         BeginProperty Column04 
            DataField       =   "caja"
            Caption         =   "   Caja"
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
            DataField       =   "valesd"
            Caption         =   "Vales Desp."
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
            DataField       =   "valese"
            Caption         =   "Vales Emp."
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
            DataField       =   "faltante"
            Caption         =   "Faltante"
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
            DataField       =   "sobrante"
            Caption         =   "Sobrante"
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
         BeginProperty Column13 
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
         BeginProperty Column14 
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
         BeginProperty Column15 
            DataField       =   "Depto7"
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
         BeginProperty Column16 
            DataField       =   "depto8"
            Caption         =   "Depto8"
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
            MarqueeStyle    =   2
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column12 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column13 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column14 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column15 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column16 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adopoliza 
         Height          =   330
         Left            =   3360
         Top             =   1320
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
         Caption         =   "Polizas"
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
   End
   Begin VB.Frame frapol 
      Height          =   735
      Left            =   120
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdPol 
         Height          =   400
         Index           =   3
         Left            =   2880
         Picture         =   "frmfacturas.frx":42AD
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Genera folios para reporte de pagos parciales y totales"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdPol 
         Height          =   400
         Index           =   2
         Left            =   1560
         Picture         =   "frmfacturas.frx":43AF
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Genera folios de pólizas"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdPol 
         Height          =   400
         Index           =   1
         Left            =   840
         Picture         =   "frmfacturas.frx":46F1
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Reporte de la póliza seleccionada"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdPol 
         Height          =   400
         Index           =   0
         Left            =   120
         Picture         =   "frmfacturas.frx":4C23
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Generar póliza de contado"
         Top             =   240
         Width           =   500
      End
   End
   Begin VB.Frame FraOpcion 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   7
         Left            =   3000
         Picture         =   "frmfacturas.frx":4F65
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Pólizas de serie contado"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   1
         Left            =   1080
         Picture         =   "frmfacturas.frx":52A7
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Abonos parciales a facturas de credito"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   4
         Left            =   2280
         Picture         =   "frmfacturas.frx":53A9
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Actualiza corte con la  factura seleccionada"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   1
         Left            =   4560
         Picture         =   "frmfacturas.frx":54AB
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Reporte de facturas para Contabilidad"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   2
         Left            =   6840
         Picture         =   "frmfacturas.frx":59DD
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Actualizar datos para reflejar cambios realizados por otros usuarios"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   3
         Left            =   5760
         Picture         =   "frmfacturas.frx":5ADF
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ir al ultimo"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   4
         Left            =   6360
         Picture         =   "frmfacturas.frx":5C51
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Busca serie y folio de la factura proporcionada"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton Command2 
         Height          =   400
         Left            =   11160
         Picture         =   "frmfacturas.frx":5DC3
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   2
         Left            =   4080
         Picture         =   "frmfacturas.frx":5F35
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Facturas Pendientes de Cobro"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   6
         Left            =   1680
         Picture         =   "frmfacturas.frx":6467
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Generación del Corte de la Serie Seleccionada"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   3
         Left            =   3600
         Picture         =   "frmfacturas.frx":65A9
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Reporte de Facturas de la serie seleccionada"
         Top             =   240
         Width           =   500
      End
      Begin VB.ComboBox cmbFiltro 
         Height          =   315
         ItemData        =   "frmfacturas.frx":6ADB
         Left            =   7440
         List            =   "frmfacturas.frx":6ADD
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   0
         Left            =   5280
         Picture         =   "frmfacturas.frx":6ADF
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ir al primero"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   5
         Left            =   600
         Picture         =   "frmfacturas.frx":6C51
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Modificar la factura seleccionada actualmente"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   0
         Left            =   120
         Picture         =   "frmfacturas.frx":6DC3
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Nueva Factura Sin Antecedentes"
         Top             =   240
         Width           =   500
      End
      Begin VB.Label lblFiltro 
         Alignment       =   2  'Center
         Caption         =   "Filtro"
         Height          =   255
         Left            =   7200
         TabIndex        =   20
         Top             =   120
         Width           =   3615
      End
   End
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   7935
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   $"frmfacturas.frx":6F05
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   9525
            MinWidth        =   9525
            Text            =   "Click en el encabezado ordena los datos en base a la columna"
            TextSave        =   "Click en el encabezado ordena los datos en base a la columna"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   600
      Top             =   6480
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
   Begin MSDataGridLib.DataGrid dbgrdVta 
      Bindings        =   "frmfacturas.frx":6F8C
      Height          =   5895
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14143944
      ColumnHeaders   =   -1  'True
      ForeColor       =   4194368
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
         DataField       =   "Noventa"
         Caption         =   "Fol. Uni."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "factura"
         Caption         =   "Factura"
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
         DataField       =   "total"
         Caption         =   "Importe"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "iva"
         Caption         =   "  Iva"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Ieps"
         Caption         =   "Ieps"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "cnombre"
         Caption         =   "                         Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "facfecha"
         Caption         =   "Fecha"
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
      BeginProperty Column07 
         DataField       =   "porpagar"
         Caption         =   "Debe"
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
         DataField       =   "cobrado"
         Caption         =   "Cobrada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "SI"
            FalseValue      =   "NO"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "faccobro"
         Caption         =   "Fecha Cobro"
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
      BeginProperty Column10 
         DataField       =   "posfechado"
         Caption         =   "Posfechado"
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
      BeginProperty Column11 
         DataField       =   "imposfechado"
         Caption         =   "$ Posfechado"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3674.835
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   1244.976
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufac 
      Caption         =   "&Facturas"
      Begin VB.Menu mnuact 
         Caption         =   "Actualizar RFC"
      End
      Begin VB.Menu mnuimp 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnucobra 
         Caption         =   "&Desactivar captura de abonos"
      End
      Begin VB.Menu mnumod 
         Caption         =   "&Cambiar de serie"
      End
      Begin VB.Menu mnurep 
         Caption         =   "&Reporte de Facturas"
      End
      Begin VB.Menu mnuz 
         Caption         =   "Captura &Observaciones"
      End
      Begin VB.Menu mnusin 
         Caption         =   "Factura &Sin Antecedentes"
      End
      Begin VB.Menu faccan 
         Caption         =   "&Cancelar sin mvto."
      End
      Begin VB.Menu mnuFacPag 
         Caption         =   "Pagos parciales y totales"
      End
      Begin VB.Menu cmbfecha 
         Caption         =   "Cambiar Fecha"
      End
      Begin VB.Menu mnumes 
         Caption         =   "Reporte de vales"
      End
      Begin VB.Menu mnufaccan 
         Caption         =   "Inserta fac. cancelada"
      End
   End
   Begin VB.Menu mnusal 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cCond As String
Private cFecha As String
Private rstCli As ADODB.Recordset
Private cUsuario As String

Private Sub adofacturas_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If AdoFacturas.Recordset!facrfc = True Then
   lblSituacion.Caption = AdoFacturas.Recordset!cnombrefac
Else
   lblSituacion.Caption = "C O N S U M I D O R   F I N A L"
End If
chkCancelado.Value = IIf(IsNull(AdoFacturas.Recordset!FACFECHACAN), 0, 1)
cmdOpcion(1).Enabled = AdoFacturas.Recordset!credito = True And AdoFacturas.Recordset!cobrado = False And IsNull(AdoFacturas.Recordset!FACFECHACAN)
End Sub


Private Sub cmbfecha_Click()
mfec = InputBox("Escriba la Fecha Correcta del Sistema", "CAMBIO FECHA")
nSepara = InStr(1, AdoFacturas.Recordset!Factura, "-")
SERIE = Mid(AdoFacturas.Recordset!Factura, 1, nSepara - 1)
Factura = Trim(Mid(AdoFacturas.Recordset!Factura, nSepara + 1, Len(AdoFacturas.Recordset!Factura)))

cn.Execute "update facventa set facfecha = '" & Trim(mfec) & "' where numfactura = '" & Trim(Factura) & "'  and serie = '" & Trim(SERIE) & "'"
cn.Execute "update facventa_det set fecha_det = '" & Trim(mfec) & "' where factura = '" & Trim(Factura) & "'  and serie = '" & Trim(SERIE) & "'"
Exit Sub
End Sub

Private Sub cmbFiltro_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub cmbFiltro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     KeyAscii = 0
     SendKeys "{TAB}"
  End If
End Sub

Private Sub cmbFiltro_LostFocus()
'MsgBox Mid(cmbFiltro.Text, (InStr(1, cmbFiltro.Text, "|") + 1))
Select Case Trim(Mid(cmbFiltro.Text, (InStr(1, cmbFiltro.Text, "|") + 1)))
Case "TOD" 'Todos
    cCond = " NOT facventa.noventa IS NULL "
    ccondrpt = "{PEDIDOS.P_pedido} <> '' "
Case "CAN" 'Canceladas
    cCond = " NOT facfechacan is null"
    ccondrpt = "{PEDIDOS.P_situacion} = 2"
Case Else
    SERIE = Trim(Mid(cmbFiltro.Text, (InStr(1, cmbFiltro.Text, "|") + 1)))
    cCond = " SERIE = '" & SERIE & "'"
End Select

cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
If chkcobro.Value = 1 Then
    cFecha = " AND Faccobro >= '" & Format(dtpFecha(0).Value, "DD-MM-YYYY") & "' and Faccobro <= '" & Format(dtpFecha(1).Value, "DD-MM-YYYY") & "'"
Else
    cFecha = " AND FacFECHA >= '" & Format(dtpFecha(0).Value, "DD-MM-YYYY") & "' and FacFECHA <= '" & Format(dtpFecha(1).Value, "DD-MM-YYYY") & "'"
End If

AdoFacturas.ConnectionString = cCadConex
AdoFacturas.CommandType = adCmdText
AdoFacturas.RecordSource = "SELECT numfactura, porpagar,faccobro,cnombrefac, Serie, facfechacan, FACVENTA.noventa, rtrim(serie) +'-'+ numfactura AS factura, facfecha, facventa.total as Total, iva, Ieps, cnombre, Facventa.cobrado, posfechado,imposfechado  FROM facventa, CatCliente WHERE  faccliente = cClave AND " & cCond & cFecha & " ORDER BY cNombre"
AdoFacturas.Refresh
For N = 0 To 4
    Cmdmoverse(N).Enabled = Not (AdoFacturas.Recordset.BOF And AdoFacturas.Recordset.EOF)
Next
lblInfo.Caption = Str(AdoFacturas.Recordset.RecordCount)
End Sub

Private Sub CmdAbonp_Click(Index As Integer)
Select Case Index
    Case 1
        Me.FraAbono.Visible = False
        Me.dbgrdVta.Enabled = True
    Case 0
        AdoAbono.Refresh
End Select
End Sub

Private Sub cmdConAceptar_Click()
Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT * FROM USUARIOS WHERE PASSMASTER = '" & Trim(txtContra.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
If RsCon.BOF And RsCon.EOF Then
   MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
   txtContra.SetFocus
   SendKeys "+{HOME}"
   Exit Sub
Else
   fraCon.Visible = False
   nOp = 10
   parciales  'Muestra frame de parcialidades
   cUsuario = RsCon!login
End If
End Sub

Private Sub cmdConCance_Click()
  Me.fraCon.Visible = False
End Sub

Private Sub cmdMoverse_Click(Index As Integer)
Dim rstBus As ADODB.Recordset
On Error GoTo Error:
If AdoFacturas.Recordset.BOF And AdoFacturas.Recordset.EOF Then
   For N = 0 To 4
       Cmdmoverse(N).Enabled = False
   Next
   Exit Sub
End If
Select Case Index
Case 0  ' Primer registro
    AdoFacturas.Recordset.MoveFirst
Case 1  ' Reporte para contabilidad
     cr1.ReportFileName = App.Path & "\Faconta.rpt"
     cr1.WindowTitle = "Reporte de facturas para Contabilidad del " & Me.dtpFecha(0).Value & " AL " & Me.dtpFecha(1).Value
     cr1.SQLQuery = "SELECT FACVENTA_DET.importe, FACVENTA_DET.tasaieps, FACVENTA_DET.factura, FACVENTA_DET.serie, FACVENTA_DET.fecha_det, " & _
                            "FACVENTA.cancelado, FACVENTA.GlobConFin, FACVENTA.rfc, " & _
                            "CATCLIENTE.cnombre, " & _
                            "FOLPOLIZA.folio, FOLPOLIZA.serie " & Chr(13) & _
                    "FROM PITICO.dbo.FACVENTA_DET FACVENTA_DET, " & _
                            "PITICO.dbo.FACVENTA FACVENTA, " & _
                            "PITICO.dbo.CATCLIENTE CATCLIENTE, " & _
                            "PITICO.dbo.FOLPOLIZA FOLPOLIZA " & Chr(13) & _
                    "WHERE FACVENTA_DET.factura = FACVENTA.numfactura AND FACVENTA_DET.serie = FACVENTA.serie AND " & _
                            "FACVENTA.faccliente = CATCLIENTE.cclave AND FACVENTA_DET.serie = '" & SERIE & "' AND " & _
                            "FACVENTA_det.fecha_det = FOLPOLIZA.fecha AND " & _
                            "FOLPOLIZA.serie = '" & SERIE & "' AND FOLPOLIZA.poliza = 1 AND " & _
                            "FACVENTA_DET.fecha_det >= '" & Format(dtpFecha(0).Value, "yyyy-dd-mm") & "' AND FACVENTA_DET.fecha_det <= '" & Format(dtpFecha(1).Value, "yyyy-dd-mm") & "' " & Chr(13) & _
                    "ORDER BY FACVENTA_DET.fecha_det ASC, FACVENTA_DET.factura ASC"
     cr1.Formulas(0) = ""
     'MsgBox cr1.SQLQuery
     cr1.Action = 1
Case 2  ' Siguiente
     Me.AdoFacturas.Refresh
Case 3  ' Ultimo
    AdoFacturas.Recordset.MoveLast
Case 4  ' Buscar clave de la venta por dia
    cCve = InputBox("Introduzca la serie y numero de factura a buscar", "Introducir factura")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    Antes = dbgrdVta.Bookmark
    AdoFacturas.Recordset.MoveFirst
      AdoFacturas.Recordset.Find "factura = '" & cCve & "'"
    If AdoFacturas.Recordset.EOF Then
       MsgBox "LA CLAVE " & UCase(cCve) & " NO SE ENCUENTRA EN LAS FACTURAS " + IIf(Me.cmbFiltro.Text = "TODAS", "" & Chr(13) & " EN EL PERIODO SELECCIONADO", cmbFiltro.Text), vbExclamation
       dbgrdVta.Bookmark = Antes
    End If
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
Dim venta As Double
Dim rs As ADODB.Recordset
Preventa = 0
Select Case Index
  Case 0 'Agregar FActura
       lprov = "SIN"
       fconfac.Show 1
       If lpprov Then
          'se agrega la factura
          Call agregarfactura(fconfac.txtafactura.Text, fconfac.txtaserie.Text, fconfac.cmbCliente.Text)
          Me.AdoFacturas.Refresh
          MsgBox "Proceso Finalizado...", vbInformation
       Else
          'MsgBox "No es posible Generar La factura si ya existe", vbInformation
       End If
  Case 1 'Cobro de la factura
       If nOp <> 10 Then  'Se me ocurrio este numero para la variable publica
          Me.fraCon.Visible = True
          Me.txtContra.SetFocus
          Exit Sub
       Else
          parciales  'Muestra frame de parcialidades
       End If
       'RESP = MsgBox("Esta seguro de Cambiar Situacion de la Factura a cobrada ? ", vbYesNo + vbQuestion, "COBRAR")
       'If RESP = vbYes Then
       '   If Len(SERIE) = 1 Then
       '      Factura = Mid(AdoFacturas.Recordset!Factura, 3, Len(AdoFacturas.Recordset!Factura))
       '   Else
       '      Factura = Mid(AdoFacturas.Recordset!Factura, 4, Len(AdoFacturas.Recordset!Factura))
       '   End If
       '   CAD = "UPDATE facventa SET cobrado = 1 , porpagar = 0, faccobro = '" & date & "' WHERE numfactura = '" & Trim(Factura) & "'  and  serie = '" & Trim(AdoFacturas.Recordset!SERIE) & "'"
       '   cn.Execute CAD
       '   MsgBox "Proceso realizado", vbInformation
       'End If
  Case 5 'Modificar numero de factura
       frmFactDet.txtcampos(2).Text = AdoFacturas.Recordset!Factura
       frmFactDet.Show 1
  Case 2
       Call REPORTEPENDIENTE
  Case 3
       Call REPORTEDECORTE
  Case 4
      venta = AdoFacturas.Recordset!noventa
      Call cortez(1, venta)
  Case 6
      Call cortez(1)
  Case 7
      Call polcontado
End Select

End Sub

'Genera polizas para series de contado
Private Sub polcontado()
FraPolizas.Visible = True
frapol.Visible = True
Adopoliza.CommandType = adCmdText
Adopoliza.ConnectionString = cCadConex
Adopoliza.RecordSource = "SELECT * FROM Polcontado ORDER BY fecha"
Adopoliza.Refresh
dbgrdpol.Caption = "VENTAS DIARIAS DEL " & Me.dtpFecha(0).Value & " AL " & Me.dtpFecha(1).Value
End Sub

Private Sub parciales()
AdoAbono.ConnectionString = cCadConex
AdoAbono.CursorType = adOpenDynamic
AdoAbono.RecordSource = "SELECT * FROM abonos WHERE rtrim(serie) +'-'+ factura = '" & Me.AdoFacturas.Recordset!Factura & "' ORDER by fecha"
AdoAbono.Refresh
FraAbono.Caption = "Abonos a la factura  [   " & AdoFacturas.Recordset!Factura & "]"
Me.FraAbono.Visible = True
Me.dbgrdVta.Enabled = False
End Sub

Private Sub cortez(tipo As Integer, Optional venta As Double)
' LA IDEA DE ESTE CORTE ES HACER UN RECORRIDO DE LAS VENTAS
'POR DIA Y PASAR LOS DATOS DE LAS VENTAS HACIA LAS FACTURAS
'EN TEORIA TODO DEBE ESTAR PERO NO PASA ASI
'1.- DE LAS TABLAS VENTAS_DET Y VENTAS BUSCAR LAS QUE NO ESTEN EN
'BUSCAR LAS FACTURAS QUE NO ESTEN EN LAS FACTURAS
'INSERTAR O ACTUALIZAR
If IsNull(Trim(SERIE)) Then
   MsgBox "Debe Seleccionar una Serie", vbInformation, "FACTURACION"
   Exit Sub
End If
If Len(SERIE) < 1 Then
   MsgBox "Debe Seleccionar una Serie", vbInformation, "FACTURACION"
   Exit Sub
End If

stb1.SimpleText = "Espere un momento; Iniciando Proceso de Facturacion..."
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
rs.CursorType = adOpenKeyset
rs.ActiveConnection = cn
Dim rsb As ADODB.Recordset
Set rsb = New ADODB.Recordset
rsb.LockType = adLockOptimistic
rsb.CursorType = adOpenKeyset
rsb.ActiveConnection = cn
fecha = dtpFecha(0).Value
RESP = MsgBox("Confirma la fecha de Generacion del Corte " & vbCrLf & Format(fecha, "Long date") & " AL " & Format(Me.dtpFecha(1).Value, "Long date"), vbYesNo + vbInformation, "CORTE ")
If RESP = vbNo Then
   Exit Sub
End If
If tipo = 2 Then
    MsgBox "Procesando por mes..." & Month(dtpFecha(0).Value), vbInformation, "MES"
    rs.Source = " select * from ventas, ventas_det where ventas.noventa =  ventas_det.noventa and month(fecha) = " & Month(fecha) & " and   Year(fecha) = " & Year(fecha) & "  and serie = '" & Trim(SERIE) & "'"
Else
    If venta > 0 Then  'El parametro de recalculo de solo una factura
       rs.Source = "Select * from ventas, ventas_det where ventas.noventa =  ventas_det.noventa AND VENTAS.NOVENTA = " & venta
    Else
       rs.Source = "SELECT * FROM ventas V, ventas_det D, facventa f WHERE v.noventa =  D.noventa AND d.factura = f.numfactura AND d.serie = f.serie AND F.facfecha >= '" & Me.dtpFecha(0).Value & "' AND facfecha <= '" & dtpFecha(1).Value & "' and f.serie = '" & Trim(SERIE) & "'"
    End If
End If
rs.Open
If rs.EOF Then
   MsgBox "No existen ventas del dia " & Format(fecha, "Long date")
   Exit Sub
Else
   rs.MoveFirst
   cn.Execute "UPDATE facventa SET faccliente = " & rs!clcliente & " WHERE numfactura = '" & rs!Factura & "' and serie = '" & rs!SERIE & "'"
   While Not rs.EOF
        stb1.SimpleText = "Procesando Venta -->  " & rs!noventa & "   " & fechafac
        'If Not Trim(rs!serie) = "B" Then
           'CUANDO ES CREDITO PROBABLEMENTE ESTE PENDIENTE POR FACTURAR"
                CAD = "select  * from facventa where numfactura = '" & Trim(rs!Factura) & "' and serie = '" & Trim(rs!SERIE) & "'"
                rsb.Source = CAD
                rsb.Open
                'If rs!noventa = 11491 Then MsgBox "wdw"
                rfc = rfccliente(rs!clcliente, IIf(rs!facrfc, 1, 0))
                If rsb.EOF Then
                    'POR DEFAULT SE PONE EL CLIENTE CONSUMIDOR FINAL
                    'ESTO NUNCA DEBE PASAR, PERO POR SI LAS DUDAS
                    'Stb1.SimpleText = "Generando Factura  " & rs!Factura
                    'DEBE SER DOS EL CLIENTE
                    CAD = "insert into facventa(faccliente,noventa,facfecha,total,iva,ieps,numfactura,serie,cobrado,rfc) values(" & _
                             2 & "," & rs!noventa & ",'" & rs!fecha & "'," & rs!total & ",0,0,'" & rs!Factura & "','" & rs!SERIE & "',1,'" & rfc & "')"
                    'MsgBox cad
                    cn.Execute CAD
                Else
                    'Rsb!facfecha = rs!fecha
                    If rsb!cancelado Then
                        rfc = "CANC999999999"
                        rsb!rfc = "CANC999999999"
                        'PARA QUE SALGA EN CEROS
                        CAD = "update facventa_det set importe = 0, rfc_det = 'CANC999999999' WHERE factura = " & Trim(rs!Factura) & " and serie = '" & Trim(rs!SERIE) & "'"
                        'MsgBox cad
                        cn.Execute CAD
                    Else
                        rsb!rfc = rfc
                    End If
                    fechafac = rsb!FACFECHA
                rsb.Update
                End If
       rsb.Close
       'If rs!tasaieps > 4 Then MsgBox "dsfsdfds"
       If IsNull(rs!tasaieps) Or rs!tasaieps < 1 Or rs!tasaieps > 8 Then
          tasaieps = tasaprod(rs!cl_producto)
       Else
          tasaieps = rs!tasaieps
       End If
       'AHORA SE BUSCA EN EL DETALLE
       rsb.Source = "select  * from facventa_det where factura = '" & Trim(rs!Factura) & "' and serie = '" & Trim(rs!SERIE) & "' AND PRODUCTO = '" & Trim(rs!cl_producto) & "'"
       'MsgBox Rsb.Source
       rsb.Open
       If rsb.EOF Then
            'SE AGREGA EL PRODUCTO
            If IsNull(rs!PREcostop) Then
                costop = 0
            Else
               costop = rs!PREcostop
            End If
            If IsNull(rs!preciop) Then
               preciop = 0
            Else
               preciop = rs!preciop
            End If
            If IsNull(rs!importe) Then
               importe = 0
            Else
               importe = rs!importe
            End If
            CAD = "INSERT INTO FACVENTA_DET(producto,cantidad,cantidadp,precio,preciop,costo,costop,importe,iva,ieps,tasaieps,serie,factura,venta,rfc_det,fecha_det) values( " & _
                       "'" & Trim(rs!cl_producto) & "'," & rs!cantidad & "," & rs!cantidadp & "," & rs!PRECIO & "," & preciop & "," & rs!PRECOSTO & "," & costop & "," & importe & "," & rs!iva & "," & rs!ieps & "," & tasaieps & ",'" & rs!SERIE & "','" & rs!Factura & "'," & rs!noventa & ",'" & rfc & "','" & fechafac & "')"
            'MsgBox cad
            cn.Execute CAD
       Else
           'CORRECCION DEL PRODUCTO EN TASAIEPS
           If Not IsNull(fechafac) Then
              rsb!fecha_det = Format(fechafac, "dd-mm-yyyy")
           End If
           rsb!tasaieps = tasaieps
           rsb!preciop = rs!preciop
           rsb!PRECIO = rs!PRECIO
           rsb!costo = rs!PRECOSTO
           rsb!costop = rs!PREcostop
           rsb!rfc_det = rfc
           rsb.Update
       End If
       rsb.Close
       'End If ' DE LA CONDICION DE LA SERIE
       rs.MoveNext
   Wend
stb1.SimpleText = " Proceso Finalizado..."
End If
End Sub

Private Function rfccliente(CLIENTE As Integer, tipo As Integer) As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
rs.CursorType = adOpenKeyset
rs.ActiveConnection = cn
rs.Open "select crfc from catcliente where cclave  = " & Trim(CLIENTE)
If rs.EOF Then
     'UPPS EXISTE UN ERROR EN LA BASE SE DEBE CORREGIR"
     rfccliente = "COOF970101111"
Else
     If tipo = 1 Then
         rfccliente = rs!crfc
     Else
         rfccliente = "COOF970101111"
    End If
End If
Set rs = Nothing
End Function

Private Function tasaprod(producto) As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
rs.CursorType = adOpenKeyset
rs.ActiveConnection = cn
rs.Open "select tasaieps from tfproduc where consec = '" & Trim(producto) & "'"
If rs.EOF Then
     'UPPS EXISTE UN ERROR EN LA BASE SE DEBE CORREGIR"
     tasaprod = 1
Else
     tasaprod = rs!tasaieps
End If
Set rs = Nothing
End Function

Private Sub REPORTEDECORTE()
On Error GoTo Error:
SERIE = Trim(Mid(cmbFiltro.Text, (InStr(1, cmbFiltro.Text, "|") + 1)))
If IsNull(Trim(SERIE)) Or Len(SERIE) < 1 Then
   MsgBox "Debe Seleccionar una Serie", vbInformation, "FACTURACION"
   Exit Sub
End If
fecha = dtpFecha(0).Value
fecha1 = dtpFecha(1).Value
cencab = "RELACION DE FACTURAS DE LA SERIE " & SERIE & " DEL " & fecha & " AL " & fecha1
ccondrpt = "{facventa_det.Fecha_DET} >= Date(" & Format(fecha, "yyyy,mm,dd") & ") and {facventa_det.Fecha_DET} <= Date (" & Format(fecha1, "yyyy,mm,dd") & ") AND {FACVENTA_det.serie} = '" & Trim(SERIE) & "'"
cmen = stb1.SimpleText
stb1.SimpleText = Space(35) & "Espere un momento, generando reporte de facturas........"
If InStr(1, "D2-G2-H2-I2-J2-DDD-GGG-JJJ-LLL-KKK", SERIE) > 0 And Mid(cSucursal, 1, 2) = 16 Then
   cr1.ReportFileName = App.Path & "\cx1.rpt"
   Tabla = "FACVTAMAY"
   Dim FechaIni As Date
   Dim FechaFin As Date
   FechaIni = "19/04/02" 'período que uilizó Cosijopi serie D2
   FechaFin = "20/05/02"
   'If (fecha >= FechaIni Or fecha1 >= FechaIni) And SERIE = "D2" Then
   If (fecha1 >= FechaIni Or fecha1 >= FechaIni) And SERIE = "D2" Then
       COND = " AND (FACVTAMAY.factura < '11100' OR (FACVTAMAY.factura > '11300' AND FACVTAMAY.factura < '11501') OR (FACVTAMAY.factura > '11600' AND FACVTAMAY.factura < '11700') OR (FACVTAMAY.factura > '11900' AND FACVTAMAY.factura < '12083') OR FACVTAMAY.factura > '12090') "
       SERIE = "D2"
   'ElseIf (fecha >= FechaIni And fecha1 <= FechaIni) And SERIE = "I2" Then
   ElseIf (fecha1 >= FechaIni And fecha1 <= FechaFin) Or (fecha <= FechaFin) And SERIE = "I2" Then
       COND = " AND ((FACVTAMAY.factura >= '11100' AND FACVTAMAY.factura <= '11300') OR (FACVTAMAY.factura >= '11501' AND FACVTAMAY.factura <= '11600') OR (FACVTAMAY.factura >= '11700' AND FACVTAMAY.factura <= '11900') OR (FACVTAMAY.factura >= '12083' AND FACVTAMAY.factura <= '12090'))  "
       SERIE = "D2"
   End If
   cadsql = "SELECT FACVTAMAY.importe, FACVTAMAY.tasaieps, FACVTAMAY.factura, FACVTAMAY.serie, FACVTAMAY.fecha_det, FACVTAMAY.rfc_det " & Chr(13) & _
            "FROM PITICO.dbo.FACVTAMAY FACVTAMAY" & Chr(13) & _
            "WHERE FACVTAMAY.serie = '" & SERIE & "' AND FACVTAMAY.fecha_det >= '" & Format(fecha, "yyyy-dd-mm") & "' AND FACVTAMAY.fecha_det <= '" & Format(fecha1, "yyyy-dd-mm") & "' " & COND & Chr(13) & _
            "ORDER BY FACVTAMAY.fecha_det ASC, FACVTAMAY.factura ASC"
Else
   cr1.ReportFileName = App.Path & "\cx.rpt"
   cadsql = "SELECT FACVENTA_DET.importe, FACVENTA_DET.tasaieps, FACVENTA_DET.factura, FACVENTA_DET.serie, FACVENTA_DET.fecha_det, FACVENTA_DET.rfc_det " & Chr(13) & _
         "FROM PITICO.dbo.FACVENTA_DET FACVENTA_DET " & Chr(13) & _
         "WHERE FACVENTA_DET.serie = '" & SERIE & "' AND FACVENTA_DET.fecha_det >= '" & Format(fecha, "yyyy-dd-mm") & "' AND  FACVENTA_DET.fecha_det <= '" & Format(fecha1, "yyyy-dd-mm") & "' " & Chr(13) & _
         "ORDER BY FACVENTA_DET.fecha_det ASC, FACVENTA_DET.factura ASC"
End If
'MsgBox cadsql
cr1.SQLQuery = cadsql
cr1.Connect = cCadConex
cr1.WindowTitle = "Reporte de Facturas cobradas"
cr1.Formulas(0) = "ENCAB = '" & cencab & "'"
cr1.Action = 1
stb1.SimpleText = cmen
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub REPORTEPENDIENTE()
On Error GoTo Error:
If IsNull(Trim(SERIE)) Or Len(SERIE) < 1 Then
   MsgBox "Debe Seleccionar una Serie", vbInformation, "FACTURACION"
   Exit Sub
End If
fecha = dtpFecha(0).Value
fecha1 = dtpFecha(1).Value
ccondrpt = "{FACVENTA.facfecha} >= Date(" & Format(fecha, "yyyy,mm,dd") & ") and {FACVENTA.facfecha} <= Date (" & Format(fecha1, "yyyy,mm,dd") & ") AND {FACVENTA.serie} = '" & Trim(SERIE) & "' AND ISNULL({FACVENTA.facfechacan}) AND {FACVENTA.cobrado} = 0"
cARcRpt = "\FacPend.rpt"
If MsgBox("DESEAS VER EL REPORTE SOLAMENTE DE LAS FACTURAS PENDIENTE DE COBRO", vbQuestion + vbYesNo, "Condición") = vbYes Then
   cencab = "FACTURAS A CREDITO PENDIENTES DE COBRO DEL DIA  " & fecha & " AL " & fecha1
   If MsgBox("REPORTE AGRUPADO POR AGENTE", vbYesNo + vbQuestion, "Tipo de reporte") = vbNo Then cARcRpt = "\FacPend1.rpt"
   ccobrado = "' AND FACVENTA.cobrado = 0 "
Else
   cencab = "FACTURAS A CREDITO DEL DIA  " & fecha & " AL " & fecha1
   ccobrado = "'"
End If
Me.stb1.SimpleText = "Espere, Generando Reporte de Facturas"
cr1.SQLQuery = "SELECT FACVENTA.facfecha, FACVENTA.numfactura, FACVENTA.serie, FACVENTA.facfechacan, FACVENTA.cobrado, FACVENTA.rfc, FACVENTA.porpagar, " & _
                       "CATCLIENTE.cclave, CATCLIENTE.cnombre, " & _
                       "VENTAS.agente " & Chr(13) & _
               "FROM pitico.dbo.FACVENTA FACVENTA, " & _
                     "PITICO.dbo.CATCLIENTE CATCLIENTE, " & _
                     "PITICO.dbo.VENTAS VENTAS " & Chr(13) & _
               "WHERE FACVENTA.faccliente = CATCLIENTE.cclave AND FACVENTA.cancelado = 0 AND " & _
                     "FACVENTA.noventa = VENTAS.noventa AND " & _
                     "FACVENTA.facfecha >= '" & Format(Me.dtpFecha(0).Value, "yyyy-dd-mm") & "' AND FACVENTA.facfecha <= '" & Format(Me.dtpFecha(1).Value, "yyyy-dd-mm") & "' AND FACVENTA.serie = '" & SERIE & ccobrado
'MsgBox cr1.SQLQuery
cr1.Connect = cCadConex
cr1.ReportFileName = App.Path & cARcRpt
cr1.WindowTitle = cencab
cr1.Formulas(0) = "ENCAB = '" & cencab & "'"
cr1.Action = 1
Exit Sub
Error:
    MsgBox Err.Description
    Unload Me
End Sub

Private Sub cmdPol_Click(Index As Integer)
Dim rs As ADODB.Recordset
Dim tmp As ADODB.Recordset
If SERIE = "" Then
   MsgBox "ES NECESARIO SELECCIONAR UNA SERIE ", vbExclamation, "Seleccione serie!"
   Exit Sub
End If
Select Case Index
    Case 0
        Set rs = New ADODB.Recordset
        Set tmp = New ADODB.Recordset
        fecha = InputBox("Fecha a generar póliza", "Fecha", date)
        If fecha = "" Then Exit Sub
        If SERIE = "" Then
           MsgBox "Es necesario especificar la serie", vbInformation, "Serie"
           Exit Sub
        End If
        rs.Open "SELECT SUM(CASE TASAIEPS WHEN 1 THEN IMPORTE END) as Depto1, " & _
                    "SUM(CASE TASAIEPS WHEN 2 THEN IMPORTE END) as Depto2, " & _
                    "SUM(CASE TASAIEPS WHEN 3 THEN IMPORTE END) as Depto3, " & _
                    "SUM(CASE TASAIEPS WHEN 4 THEN IMPORTE END) as Depto4, " & _
                    "SUM(CASE TASAIEPS WHEN 5 THEN IMPORTE END) as Depto5, " & _
                    "SUM(CASE TASAIEPS WHEN 6 THEN IMPORTE END) as Depto6, " & _
                    "SUM(CASE TASAIEPS WHEN 7 THEN IMPORTE END) as Depto7, " & _
                    "SUM(CASE TASAIEPS WHEN 8 THEN IMPORTE END) as Depto8 " & _
                "FROM facventa_det WHERE fecha_det >= '" & fecha & "' And fecha_det <= '" & fecha & "' " & _
                    "And serie = '" & SERIE & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
        tmp.Open "SELECT * FROM polcontado WHERE fecha = '" & fecha & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
        If tmp.BOF And tmp.EOF Then
           tmp.AddNew
           tmp!Folio = 0
           tmp!fecha = fecha
           tmp!SERIE = SERIE
           tmp!bancos = 0
           tmp!efectivo = 0
           tmp!valesD = 0
           tmp!valesE = 0
           tmp!Faltante = 0
           tmp!Sobrante = 0
           tmp!sucursal = Trim(Mid(cSucursal, 1, 3))
        End If
        tmp!depto1 = IIf(IsNull(rs!depto1), 0, rs!depto1)
        tmp!depto2 = IIf(IsNull(rs!depto2), 0, rs!depto2)
        tmp!Iva2 = IIf(IsNull(rs!depto2), 0, rs!depto2 - (rs!depto2 / 1.15))
        tmp!depto3 = IIf(IsNull(rs!depto3), 0, rs!depto3)
        tmp!Iva3 = IIf(IsNull(rs!depto3), 0, rs!depto3 - (rs!depto3 / 1.15))
        tmp!depto4 = IIf(IsNull(rs!depto4), 0, rs!depto4)
        tmp!Iva4 = IIf(IsNull(rs!depto4), 0, rs!depto4 - (rs!depto4 / 1.15))
        tmp!depto5 = IIf(IsNull(rs!depto5), 0, rs!depto5)
        tmp!Iva5 = IIf(IsNull(rs!depto5), 0, rs!depto5 - (rs!depto5 / 1.15))
        tmp!depto6 = IIf(IsNull(rs!depto6), 0, rs!depto6)
        tmp!Iva6 = IIf(IsNull(rs!depto6), 0, rs!depto6 - (rs!depto6 / 1.15))
        tmp!depto7 = IIf(IsNull(rs!depto7), 0, rs!depto7)
        tmp!Iva7 = IIf(IsNull(rs!depto7), 0, rs!depto7 - (rs!depto7 / 1.15))
        tmp!depto8 = IIf(IsNull(rs!depto8), 0, rs!depto8)
        tmp!Iva8 = IIf(IsNull(rs!depto8), 0, rs!depto8 - (rs!depto8 / 1.15))
        tmp.Update
        Adopoliza.Refresh
    Case 1
         cr1.ReportFileName = App.Path & "\Polconta.rpt"
         cr1.Connect = cCadConex
         cr1.WindowTitle = "POLIZA DE LA SERIE " & SERIE & " DE FECHA " & Me.Adopoliza.Recordset!fecha
         cr1.Formulas(0) = "FOLIO = " & Adopoliza.Recordset!Folio
         cr1.Formulas(1) = "ENCAB = 'POLIZA CONTADO DE LA SERIE " & SERIE & " DE FECHA " & Adopoliza.Recordset!fecha & "'"
         'MsgBox cr1.Formulas(1)
         cr1.Action = 1
    Case 2  'Folios para polizas de ingreso
        Dim fechaP As Date
        Dim fecmax As Date
        Select Case SERIE
             Case "Y1"
                 folini = 2100
             Case "B"
                 folini = 2800
             Case "D2"
                 folini = 2200
             Case "I2"
                 folini = 2300
             Case "G2"
                 folini = 4600
             Case "H2"
                 folini = 4500
             Case "GGG"
                 folini = 5300
             Case "HHH"
                 folini = 5400
             Case "D"
                 folini = 5700
             Case "AB"
                 folini = 6000
        End Select
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM folpoliza WHERE serie = '" & SERIE & "' AND  poliza = 1 ORDER BY fecha", cn, adOpenKeyset, , adCmdText
        If rs.BOF And rs.EOF Then
           N = folini
           MESANT = Month(dtpFecha(0).Value)
           fechaP = dtpFecha(0) - 1
        Else
           rs.MoveLast
           MESANT = Month(rs!fecha)
           fechaP = rs!fecha
           N = rs!Folio
        End If
        rs.Close
        While fechaP <= dtpFecha(1).Value
            fechaP = fechaP + 1
            rs.Open "SELECT DISTINCT FACFECHA FROM FACVENTA WHERE FACFECHA = '" & fechaP & "' AND SERIE = '" & SERIE & "'", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rs.BOF And rs.EOF) Then
               If MESANT <> Month(rs!FACFECHA) Then N = folini '2101 => Primer folio del mes
               N = N + 1
               cn.Execute "INSERT INTO FOLPOLIZA VALUES ('" & rs!FACFECHA & "'," & N & ",'" & SERIE & "',1)"
               MESANT = Month(rs!FACFECHA)
            End If
            rs.Close
        Wend
        MsgBox "LOS FOLIOS SE GENERARON HASTA EL DIA " & Me.dtpFecha(1).Value, vbInformation, "Folios"
    Case 3   'Folios para abonos de facturas
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM folpoliza WHERE serie = '" & SERIE & "' and poliza = 0 ORDER BY fecha", cn, adOpenKeyset, , adCmdText
        If rs.BOF And rs.EOF Then
           N = 0
           fechaP = "01/04/2003"
        Else
           rs.MoveLast
           fechaP = rs!fecha
           N = rs!Folio
        End If
        rs.Close
        While fechaP <= dtpFecha(1).Value
            fechaP = fechaP + 1
            rs.Open "SELECT DISTINCT fecha FROM abonos WHERE FECHA = '" & fechaP & "' AND SERIE = '" & SERIE & "' AND posfechado IS NULL", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rs.BOF And rs.EOF) Then
               N = N + 1
               cn.Execute "INSERT INTO FOLPOLIZA VALUES ('" & rs!fecha & "'," & N & ",'" & SERIE & "',0)"
            End If
            rs.Close
        Wend
        MsgBox "LOS FOLIOS PARA RELACION DE ABONOS SE GENERO CORRECTAMENTE", vbInformation, "Folios Abonos"
End Select
End Sub

Private Sub Command2_Click()
If FraPolizas.Visible = False Then
   Unload Me
   frmAreaRecibo.Show
Else
   frapol.Visible = False
   FraPolizas.Visible = False
End If
End Sub

Private Sub dbgrdAbono_AfterUpdate()
On Error Resume Next
Dim rs As ADODB.Recordset
Dim TOTABONO
 Set rs = New ADODB.Recordset
 rs.Open "SELECT SUM(importe) as TotAbono FROM abonos WHERE posfechado IS null and RTRIM(serie)+ '-'+ factura = '" & Me.AdoFacturas.Recordset!Factura & "' AND fecha <= '" & Format(dbgrdAbono.Columns(5).Value, "dd-mm-yyyy") & "'", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
 TOTABONO = IIf(IsNull(rs!TOTABONO), 0, rs!TOTABONO)
 AdoAbono.Recordset!TotFac = AdoFacturas.Recordset!total
 AdoAbono.Recordset!TOTABONO = TOTABONO
 AdoAbono.Recordset!debe = AdoAbono.Recordset!TotFac - TOTABONO
 AdoAbono.Recordset!USUARIO = Trim(cUsuario)
 AdoAbono.Recordset.Update
 rs.Close
 Set rs = Nothing
End Sub

Private Sub dbgrdAbono_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim RST As ADODB.Recordset
Set RST = New ADODB.Recordset
RST.Open "SELECT * FROM diasver WHERE fechaveri = '" & dbgrdAbono.Columns(5).Value & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not (RST.BOF And RST.EOF) Then
   MsgBox "NO ES POSIBLE MODIFICAR ESTE FECHA PORQUE YA FUE VALIDADA", vbInformation, "Fecha Verificada"
   Cancel = True
   AdoAbono.Refresh
ElseIf UCase(dbgrdAbono.Columns(ColIndex).DataField) = "POSFECHADO" Then
   If Trim(dbgrdAbono.Columns(ColIndex).Text) = "" Then
      cn.Execute "UPDATE facventa SET posfechado = NULL, imposfechado = NULL WHERE numfactura = '" & AdoAbono.Recordset!Factura & "' AND serie  = '" & AdoAbono.Recordset!SERIE & "'"
   Else
     ' cn.Execute "UPDATE facventa SET posfechado = '" & dbgrdAbono.Columns(ColIndex).Text & "', imposfechado = " & AdoAbono.Recordset!importe & " WHERE numfactura = '" & AdoAbono.Recordset!Factura & "' AND serie  = '" & AdoAbono.Recordset!SERIE & "'"
      'MsgBox "UPDATE facventa SET posfechado = '" & dbgrdAbono.Columns(ColIndex).Text & "', imposfechado = (SELECT SUM(importe) FROM abonos where serie = '" & AdoAbono.Recordset!SERIE & "' and factura = " & AdoAbono.Recordset!Factura & " and not posfechado is null ) WHERE numfactura = " & AdoAbono.Recordset!Factura & "' AND serie  = '" & AdoAbono.Recordset!SERIE & "'"
      cn.Execute "UPDATE facventa SET posfechado = '" & dbgrdAbono.Columns(ColIndex).Text & "', imposfechado = (SELECT SUM(importe) FROM abonos where serie = '" & AdoAbono.Recordset!SERIE & "' and factura = " & AdoAbono.Recordset!Factura & " and not posfechado is null ) WHERE numfactura = " & AdoAbono.Recordset!Factura & " AND serie  = '" & AdoAbono.Recordset!SERIE & "'"
   End If
End If
End Sub

Private Sub dbgrdAbono_GotFocus()
If AdoAbono.Recordset.BOF And AdoAbono.Recordset.EOF Then
   AdoAbono.Recordset.AddNew
   AdoAbono.Recordset!debe = Me.AdoFacturas.Recordset!total
   Call iniciaVals
End If
End Sub

Private Sub dbgrdAbono_LostFocus()
On Error GoTo Error:
Dim rs As ADODB.Recordset
Dim TOTABONO
 Set rs = New ADODB.Recordset
 rs.Open "SELECT SUM(importe) as TotAbono FROM abonos WHERE posfechado IS NULL and RTRIM(serie)+ '-'+ factura = '" & Me.AdoFacturas.Recordset!Factura & "'", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
 TOTABONO = IIf(IsNull(rs!TOTABONO), 0, rs!TOTABONO)
 If TOTABONO >= AdoFacturas.Recordset!total Then
    cn.Execute "UPDATE facventa SET cobrado = 1 , porpagar = 0, faccobro = '" & date & "' WHERE RTRIM(serie)+ '-'+ NUMfactura = '" & Me.AdoFacturas.Recordset!Factura & "'"
 Else
    cn.Execute "UPDATE facventa SET cobrado = 0 , porpagar = " & AdoFacturas.Recordset!total - TOTABONO & " WHERE RTRIM(serie)+ '-'+ NUMfactura = '" & Me.AdoFacturas.Recordset!Factura & "'"
 End If
 rs.Close
 Set rs = Nothing
 Exit Sub
Error:
    MsgBox "ERROR AL ACTUALIZAR LOS ABONOS DE LA FACTURA", vbCritical
End Sub

Private Sub dbgrdAbono_OnAddNew()
   Call iniciaVals
End Sub

Private Sub iniciaVals()
   dbgrdAbono.Columns(0).Value = Trim(AdoFacturas.Recordset!SERIE)
   dbgrdAbono.Columns(1).Value = AdoFacturas.Recordset!numfactura
   dbgrdAbono.Columns(2).Value = 0
   dbgrdAbono.Columns(3).Value = 0
   'dbgrdAbono.Columns(4).Value = 0
   dbgrdAbono.Columns(5).Value = date
   dbgrdAbono.Columns(6).Value = 0
   dbgrdAbono.Columns(7).Value = "EFECTIVO"
   dbgrdAbono.Columns(8).Value = 0
End Sub

Private Sub dbgrdpol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
End If
End Sub

Private Sub dbgrdVta_DblClick()
  cmdOpcion_Click 5
End Sub

Private Sub dbgrdVta_HeadClick(ByVal ColIndex As Integer)
  stb1.SimpleText = Space(65) + "Espere un momento ordenando Pedidos por " & dbgrdVta.Columns(ColIndex).Caption
  If UCase(dbgrdVta.Columns(ColIndex).DataField) = "CNOMBRE" Then
     AdoFacturas.RecordSource = "SELECT porpagar,faccobro,cnombrefac, Serie, NUMFACTURA,facfechacan, FACVENTA.noventa, rtrim(serie) +'-'+ numfactura AS factura, facfecha, facventa.total as Total, iva, Ieps, cnombre, Facventa.cobrado, posfechado,imposfechado  FROM facventa, CatCliente WHERE  faccliente = cClave AND " & cCond & cFecha & " ORDER BY " & dbgrdVta.Columns(ColIndex).DataField
  Else
     AdoFacturas.RecordSource = "SELECT porpagar,faccobro,cnombrefac, Serie, numfactura,facfechacan, FACVENTA.noventa, rtrim(serie) +'-'+ numfactura AS factura, facfecha, facventa.total as Total, iva, Ieps, cnombre, Facventa.cobrado, posfechado,imposfechado  FROM facventa, CatCliente WHERE  faccliente = cClave AND " & cCond & cFecha & " ORDER BY facventa." & dbgrdVta.Columns(ColIndex).DataField
  End If
  AdoFacturas.Refresh
  stb1.Panels(1).Text = Space(85) + "Ventas ordenandas por " & dbgrdVta.Columns(ColIndex).Caption
End Sub

Private Sub dbgrdVta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdOpcion_Click 5
End Sub

Private Sub dtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtpFecha_LostFocus(Index As Integer)
On Error GoTo Error:
 
'cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
If chkcobro.Value = 1 Then
    cFecha = " AND Faccobro > = '" & dtpFecha(0).Value & "' and  Faccobro <= '" & dtpFecha(1).Value & "'"
Else
    cFecha = " AND Facfecha > = '" & dtpFecha(0).Value & "' and  Facfecha <= '" & dtpFecha(1).Value & "'"
End If
 AdoFacturas.RecordSource = "SELECT numfactura, porpagar,faccobro,cnombrefac, Serie,NUMFACTURA, facfechacan, FACVENTA.noventa, rtrim(serie) +'-'+ numfactura AS factura, facfecha, facventa.total as Total, iva, Ieps, cnombre, Facventa.cobrado, posfechado,imposfechado  FROM facventa, CatCliente WHERE  faccliente = cClave AND " & cCond & cFecha & " ORDER BY cNombre"
 AdoFacturas.Refresh
' MsgBox AdoFacturas.RecordSource
 lblInfo.Caption = Str(AdoFacturas.Recordset.RecordCount)
 For N = 0 To 4   'Si esta vacio el recordset desactivo las opciones
   Cmdmoverse(N).Enabled = Not (AdoFacturas.Recordset.BOF And AdoFacturas.Recordset.EOF)
 Next
 Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub faccan_Click()
SERIE = Mid(AdoFacturas.Recordset!Factura, 1, InStr(1, AdoFacturas.Recordset!Factura, "-") - 1)
If Mid(SERIE, 2, 1) = "-" Then
   SERIE = Mid(SERIE, 1, 1)
End If
If Len(SERIE) = 1 Then
   Factura = Mid(AdoFacturas.Recordset!Factura, InStr(1, AdoFacturas.Recordset!Factura, "-") + 1, Len(AdoFacturas.Recordset!Factura))
Else
   Factura = Mid(AdoFacturas.Recordset!Factura, InStr(1, AdoFacturas.Recordset!Factura, "-") + 1, Len(AdoFacturas.Recordset!Factura))
End If
cn.Execute "UPDATE facventa SET facfechacan = '" & date & " " & Time & "', cancelado = 1, TOTAL = 0  WHERE numfactura = '" & Trim(Factura) & "' AND serie = '" & Trim(SERIE) & "'"
cn.Execute "UPDATE facventa_det SET rfc_det = 'CANC999999999',IMPORTE = 0  WHERE factura = '" & Trim(Factura) & "'  and serie = '" & Trim(SERIE) & "'"
cn.Execute "UPDATE ventas_det SET cancelado = 1, importe = 0 WHERE factura = '" & Trim(Factura) & "'  and serie = '" & Trim(SERIE) & "'"
MsgBox "SE ACTUALIZO SOLAMENTE LA FACTURA COMO CANCELADA, NO SE INCREMENTA INVENTARIO Y TAMPOCO SE GENERA VENTA", vbInformation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 119
     frmCalc.Show 1
  Case 120
        Me.fraCon.Visible = True
        Me.txtContra.Text = ""
        Me.txtContra.SetFocus
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Form_Load()
  cmbFiltro.AddItem "TODAS       | TOD"
  cmbFiltro.AddItem "CANCELADAS      | CAN"
  cmbFiltro.AddItem "----------------------------------------------------------------------"
  cmbFiltro.AddItem "MIGUEL CABRERA CONTADO      | Y1"
  cmbFiltro.AddItem "MIGUEL CABRERA PREVENTA     | CCC"
  cmbFiltro.AddItem "MIGUEL CABRERA CREDITO      | B"
  cmbFiltro.AddItem "CENTRAL MODULO N | D2"
  cmbFiltro.AddItem "PUERTO ESCONDIDO CONTADO       | G2"
  cmbFiltro.AddItem "PUERTO ESCONDIDO CREDITO       | H2"
  cmbFiltro.AddItem "PUERTO ESCONDIDO PREVENTA      | DDD"
  cmbFiltro.AddItem "COSIJOPI CENTRAL CONTADO       | I2"
  cmbFiltro.AddItem "COSIJOPI CENTRAL CREDITO       | J2"
  cmbFiltro.AddItem "MIAHUATLAN CONTADO       | GGG"
  cmbFiltro.AddItem "MIAHUATLAN CREDITO       | HHH"
  cmbFiltro.AddItem "ISTMO CONTADO ANTERIOR   | D"
  cmbFiltro.AddItem "ISTMO CONTADO NUEVO      | JJJ"
  cmbFiltro.AddItem "ISTMO PREVENTA           | LLL"
  cmbFiltro.AddItem "ISTMO CREDITO            | KKK"
  cmbFiltro.AddItem "TAPACHULA CONTADO-PIT13  | AB"
  cmbFiltro.AddItem "TAPACHULA CREDITO-PIT13  | ABX"

  cmbFiltro.ListIndex = 0
  cCond = " NOT facventa.NOVENTA IS NULL "
  If dtpFecha(0).Value = "01/01/01" Then dtpFecha(0).Value = Format(date, "DD/mm/yyyy")
  If dtpFecha(1).Value = "01/01/01" Then dtpFecha(1).Value = Format(date, "DD/mm/yyyy")
  
  cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
  cFecha = " AND facfecha >= '" & dtpFecha(0).Value & "' and facfecha <= '" & dtpFecha(0).Value & "'"   'Cargo todas las ventas
  AdoFacturas.ConnectionString = cn
  AdoFacturas.CommandType = adCmdText
  AdoFacturas.RecordSource = "SELECT porpagar,faccobro,cnombrefac, Serie, NUMFACTURA,facfechacan, posfechado,imposfechado, FACVENTA.noventa, rtrim(serie) +'-'+ numfactura AS factura, facfecha, facventa.total as Total, iva, Ieps, cnombre, Facventa.cobrado, Facventa.faccobro FROM facventa, CatCliente WHERE  faccliente = cClave AND " & cCond & cFecha & " ORDER BY cNombre"
  AdoFacturas.Refresh
  For N = 0 To 4
     Cmdmoverse(N).Enabled = Not (AdoFacturas.Recordset.BOF And AdoFacturas.Recordset.EOF)
  Next
  lblInfo.Caption = Str(AdoFacturas.Recordset.RecordCount)
  nOp = 0
End Sub

Private Sub FACTURAIMPRIMIRv(Factura As String, SERIE As String)
' On Error GoTo Error:
Dim rstDetVta As ADODB.Recordset
Dim rstDir As ADODB.Recordset
Dim Impresora As Printer
Dim nTotal
Dim nCajas
Dim lNvaFac As Boolean
Dim nAlto
Dim rsFac As ADODB.Recordset
Dim TotFac As Double
Dim LETRAS
lNvaFac = True: nProd = 0
nAlto = -1.1
'ProdenFac = 14
'CONFIGURACION DE LA IMPRESION
Printer.ScaleMode = vbCentimeters
Printer.FontName = "ARIAL NARROW"
Printer.FontSize = 8
Printer.Width = 12190
Printer.Height = 7938
Printer.PaperSize = 1
Printer.Orientation = 2
'EL ENCABEZADO

Set rsFac = New ADODB.Recordset
'cad = "select cciudad,ccolonia, cdireccion, cnombre, faccliente,facfecha,total,rfc FROM facventa F, catcliente WHERE cclave = faccliente  and NUMFACTURA =  '" & Trim(factura) & "' AND SERIE = '" & Trim(serie) & "'"
CAD = "select cciudad,ccolonia, cdireccion, cnombrefac cnombre, faccliente,facfecha,total,rfc FROM facventa F, catcliente WHERE cclave = faccliente  and NUMFACTURA =  '" & Trim(Factura) & "' AND SERIE = '" & Trim(SERIE) & "'"
rsFac.Open CAD, cn, adOpenKeyset, adLockOptimistic, 1
'cad = "select cciudad,ccolonia, cdireccion, cnombre, faccliente,facfecha,total,rfc,producto,cantidad,cantidadp,preciop,precio,costo,costop,d.iva, d.ieps, D.IMPORTE, d.tasaieps,d.factura,d.serie,descripc, contenid,medida, paquetes FROM facventa F, facventa_det D, TFPRODUC  T , catcliente WHERE cclave = faccliente  and NUMFACTURA = FACTURA AND F.SERIE = D.SERIE AND D.PRODUCTO = T.CONSEC AND FACTURA = '" & Trim(factura) & "' AND d.SERIE = '" & Trim(serie) & "'"
'rsFac.Open cad, cn, adOpenDynamic, adLockOptimistic, adCmdText
If IsNull(rsFac!rfc) Then
    MsgBox "Es necesario Especificar un rfc para la factura" & vbCrLf & "Por Favor Escriba el rfc Correcto", vbInformation
    RESP = InputBox("Escriba el RFC correcto", "RFC", "COOF970101111")
    rfc = RESP '"COOF970101111"
    CAD = "UPDATE FACVENTA SET RFC =  '" & Trim(rfc) & "'   WHERE NUMFACTURA = '" & Trim(Factura) & "'  AND SERIE = '" & Trim(SERIE) & "'"
    cn.Execute CAD
Else
    rfc = rsFac!rfc
End If

Printer.CurrentY = 3.8 + nAlto
Printer.CurrentX = 0.5
Printer.Print IIf(rfc = "COOF970101111", "C O N S U M I D O R   F I N A L", rsFac!cNombre);
Printer.CurrentX = 10
Printer.Print rfc
Printer.CurrentX = 0.5
Printer.Print IIf(rsFac!rfc = "COOF970101111", "C O N O C I D O", rsFac!cdireccion);
Printer.CurrentX = 10
Printer.Print " " 'rscli!cTelefono
Printer.CurrentX = 0.5
Printer.Print rsFac!ccolonia;
Printer.CurrentX = 10
Printer.Print rsFac!cciudad
Printer.CurrentX = 0.5
Printer.Print Trim(SERIE) & "  " & Factura;
Printer.CurrentY = 5 + nAlto
'Printer.CurrentX = 16.5
Printer.Print Format(rsFac!FACFECHA, "long date") & Space(3) & Format(Time, "HH:MM AM/PM")
NVENTA1 = 0: NIVA1 = 0: NIEPS1 = 0
NVENTA2 = 0: NIVA2 = 0: NIEPS2 = 0
NVENTA3 = 0: NIVA3 = 0: NIEPS3 = 0
NVENTA4 = 0: NIVA4 = 0: NIEPS4 = 0
NVENTA5 = 0: NIVA5 = 0: NIEPS5 = 0
NVENTA6 = 0: NIVA6 = 0: NIEPS6 = 0
NVENTA7 = 0: NIVA7 = 0: NIEPS7 = 0
Printer.CurrentY = 6.5 + nAlto
rsFac.Close
'EL DETALLE
CAD = "SELECT producto,cantidad,cantidadp,preciop,precio,costo,costop,d.iva, d.ieps, D.IMPORTE, d.tasaieps,d.factura,d.serie,descripc, LTRIM(STR(T.paquetes)) + ' X ' +  lTrim(str(T.contenid,10,3)) + space(2) + SUBSTRING(T.medida,1,5)  AS medida , paquetes FROM facventa_det D, TFPRODUC  T  WHERE D.PRODUCTO = T.CONSEC AND FACTURA = '" & Trim(Factura) & "' AND d.SERIE = '" & Trim(SERIE) & "'"
rsFac.Open CAD, cn, adOpenKeyset, adLockOptimistic, 1
rsFac.MoveFirst
While Not rsFac.EOF
    Printer.CurrentX = 0.3
    If rsFac!cantidad > 0 Then
       Printer.Print rsFac!cantidad & "CJ" & IIf(rsFac!cantidadp > 0, "-" & rsFac!cantidadp & "PZ", "");
    Else
       Printer.CurrentX = 0.3
       Printer.Print rsFac!cantidadp & "PZ";
    End If
    
    Printer.CurrentX = 1.5
    If Printer.TextWidth(rsFac!descripc) > 5.5 Then
       For N = 1 To Len(rsFac!descripc)
          If Printer.TextWidth(Mid(rsFac!descripc, 1, N)) > 5 Then Exit For
       Next
       Printer.Print Mid(rsFac!descripc, 1, N);
    Else
       Printer.Print rsFac!descripc;
    End If
    
    Printer.CurrentX = 6.5
    Printer.Print rsFac!medida;
    Printer.CurrentX = 9.5
    Printer.Print Format(rsFac!ieps, "00");
    Printer.CurrentX = 10
    Printer.Print Format(rsFac!iva, "00");
    Printer.CurrentX = 10.3
    Printer.Print String(10 - Len(Trim(Format(rsFac!PRECIO, "########0.00"))), " ") & Format(rsFac!PRECIO, "########0.00");
    Printer.CurrentX = 11.8
    Printer.Print String(12 - Len(Trim(Format(rsFac!importe, "########0.00"))), " ") & Format(rsFac!importe, "########0.00")
    ncosto = rsFac!importe
    If rsFac!iva = 0 And rsFac!ieps = 0 Then        'Depto 1
       NVENTA1 = NVENTA1 + ncosto
       NIVA1 = 0
       NIEPS1 = 0
    ElseIf rsFac!iva = 15 And rsFac!ieps = 0 Then   'Depto 2
       NVENTA2 = NVENTA2 + ncosto
       NIVA2 = NIVA2 + (ncosto / 1.15 * (15 / 100))
       NIEPS2 = 0
    ElseIf rsFac!iva = 15 And rsFac!ieps = 25 Then  'Depto 3
       NVENTA3 = NVENTA3 + ncosto
       NIVA3 = NIVA3 + (ncosto / 1.15 * (15 / 100))
       NIEPS3 = NIEPS3 + (((ncosto / 1.15) / 1.25) * 25 / 100)
    ElseIf rsFac!iva = 15 And rsFac!ieps = 30 Then 'Depto 4
       NVENTA4 = NVENTA4 + ncosto
       NIVA4 = NIVA4 + (ncosto / 1.15 * (15 / 100))
       NIEPS4 = NIEPS4 + (((ncosto / 1.15) / 1.3) * 30 / 100)
    ElseIf rsFac!iva = 15 And rsFac!ieps = 50 Then 'Depto 5
       NVENTA5 = NVENTA5 + ncosto
       NIVA5 = NIVA5 + (ncosto / 1.15 * (15 / 100))
       NIEPS5 = NIEPS5 + (((ncosto / 1.15) / 1.5) * 50 / 100)
    ElseIf rsFac!iva = 15 And rsFac!ieps = 60 Then 'Depto 6
       NVENTA6 = NVENTA6 + ncosto
       NIVA6 = NIVA6 + (ncosto / 1.15 * (15 / 100))
       NIEPS6 = NIEPS6 + (((ncosto / 1.15) / 1.6) * 60 / 100)
    ElseIf rsFac!iva = 15 And rsFac!ieps = 5 Then 'Depto 7
       NVENTA7 = NVENTA7 + ncosto
       NIVA7 = NIVA7 + (ncosto / 1.15 * (15 / 100))
       NIEPS7 = NIEPS7 + (((ncosto / 1.15) / 1.05) * 5 / 100)
    End If
    nProd = nProd + 1
    rsFac.MoveNext
Wend
Printer.CurrentY = 18.5 + nAlto
Printer.CurrentX = 1
Printer.Print "DEP1";
Printer.CurrentX = 2
Printer.Print "DEP2";
Printer.CurrentX = 3
Printer.Print "DEP3";
Printer.CurrentX = 4
Printer.Print "DEP4";
'Printer.CurrentY = 16.5 + nAlto
Printer.CurrentX = 5
Printer.Print "DEP5";
'Printer.CurrentY = 16.5 + nAlto
Printer.CurrentX = 6
Printer.Print "DEP6";
'Printer.CurrentY = 16.5 + nAlto
Printer.CurrentX = 7
Printer.Print "DEP7"
'SUBTOTALES
'Cuando es consumidor final no se desglosa la factura ni se imprime ieps e iva

Printer.CurrentX = 1
Printer.Print Format(NVENTA1, "########0.00");
Printer.CurrentX = 2
'En el caso de que sea consumidor final no se presenta el ieps y iva
Printer.Print Format(NVENTA2, "########0.00");
Printer.CurrentX = 3
Printer.Print Format(NVENTA3, "########0.00");
Printer.CurrentX = 4
Printer.Print Format(NVENTA4, "########0.00");
Printer.CurrentX = 5
Printer.Print Format(NVENTA5, "########0.00");
Printer.CurrentX = 6
Printer.Print Format(NVENTA6, "########0.00");
Printer.CurrentX = 7
Printer.Print Format(NVENTA7, "########0.00");

Printer.CurrentX = 10
Printer.Print "SUBTOTAL";
Printer.CurrentX = 11.5
Printer.FontSize = 10
Printer.Print String(11 - Len(Trim(Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4) + (NVENTA5 - NIVA5 - NIEPS5) + (NVENTA6 - NIVA6 - NIEPS6) + (NVENTA7 - NIVA7 - NIEPS7), "#########0.00"))), " ") & Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4), "#########0.00")
Printer.FontSize = 8
'IVA
Printer.CurrentX = 1
'CONDICIONES CUANDO NO ES CONSUMIDOR FINAL
Printer.Print Format(NIVA1, "########0.00");
Printer.CurrentX = 2
Printer.Print Format(NIVA2, "########0.00");
Printer.CurrentX = 3
Printer.Print Format(NIVA3, "########0.00");
Printer.CurrentX = 4
Printer.Print Format(NIVA4, "########0.00");
Printer.CurrentX = 5
Printer.Print Format(NIVA5, "########0.00");
Printer.CurrentX = 6
Printer.Print Format(NIVA6, "########0.00");
Printer.CurrentX = 7
Printer.Print Format(NIVA7, "########0.00");
Printer.CurrentX = 10
Printer.Print "IVA";
Printer.CurrentX = 11.5
Printer.FontSize = 10
Printer.Print String(10 - Len(Trim(Format(NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7, "#########0.00"))), " ") & Format(NIVA2 + NIVA3 + NIVA4, "#########0.00")
Printer.FontSize = 8
'IEPS
Printer.CurrentX = 1
Printer.Print Format(NIEPS1, "########0.00");
Printer.CurrentX = 2
Printer.Print Format(NIEPS2, "########0.00");
Printer.CurrentX = 3
Printer.Print Format(NIEPS3, "########0.00");
Printer.CurrentX = 4
Printer.Print Format(NIEPS4, "########0.00");
Printer.CurrentX = 5
Printer.Print Format(NIEPS5, "########0.00");
Printer.CurrentX = 6
Printer.Print Format(NIEPS6, "########0.00");
Printer.CurrentX = 7
Printer.Print Format(NIEPS7, "########0.00");
Printer.CurrentX = 10
Printer.Print "IEPS";
Printer.CurrentX = 11.5
Printer.FontSize = 10
Printer.Print String(10 - Len(Trim(Format(NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7, "#########0.00"))), " ") & Format(NIEPS3 + NIEPS4, "#########0.00")
Printer.FontSize = 8

Printer.FontSize = 8
'TOTAL DE LA VENTA
Printer.CurrentX = 1
If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
Printer.CurrentX = 2
If NVENTA2 > 0 Then Printer.Print Format(NVENTA2, "########0.00");
Printer.CurrentX = 3
If NVENTA3 > 0 Then Printer.Print Format(NVENTA3, "########0.00");
Printer.CurrentX = 4
If NVENTA4 > 0 Then Printer.Print Format(NVENTA4, "########0.00");
Printer.CurrentX = 5
If NVENTA5 > 0 Then Printer.Print Format(NVENTA5, "########0.00");
Printer.CurrentX = 6
If NVENTA6 > 0 Then Printer.Print Format(NVENTA6, "########0.00");
Printer.CurrentX = 7
If NVENTA7 > 0 Then Printer.Print Format(NVENTA7, "########0.00");

Printer.CurrentX = 10
Printer.Print "TOTAL";
Printer.CurrentX = 11.5
Printer.FontSize = 10
Printer.Print String(11 - Len(Trim(Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7, "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7, "#########0.00")
Printer.FontSize = 8
Printer.CurrentX = 2
TotFac = NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7
LETRAS = frmVentas.NumLet$(TotFac)
Printer.Print LETRAS
lNvaFac = True: nProd = 0
Printer.EndDoc
MsgBox "LA IMPRESION SE REALIZO CORRECTAMENTE", vbInformation, "Impresión finalizada"
Exit Sub
Error:
  MsgBox Err.Description, vbInformation
End Sub


Private Sub FACTURAIMPRIMIRh(Factura As String, SERIE As String)
' On Error GoTo Error:
Dim rstDetVta As ADODB.Recordset
Dim rstDir As ADODB.Recordset
Dim Impresora As Printer
Dim nTotal
Dim nCajas
Dim lNvaFac As Boolean
Dim nAlto
Dim rsFac As ADODB.Recordset
Dim TotFac As Double
Dim LETRAS
lNvaFac = True: nProd = 0
nAlto = -1.1
'ProdenFac = 14
'CONFIGURACION DE LA IMPRESION
Printer.ScaleMode = vbCentimeters
Printer.FontName = "ARIAL NARROW"
Printer.FontSize = 8
Printer.Width = 12190
Printer.Height = 7938
Printer.PaperSize = 1
Printer.Orientation = 1
'EL ENCABEZADO

Set rsFac = New ADODB.Recordset
'cad = "select cciudad,ccolonia, cdireccion, cnombre, faccliente,facfecha,total,rfc FROM facventa F, catcliente WHERE cclave = faccliente  and NUMFACTURA =  '" & Trim(factura) & "' AND SERIE = '" & Trim(serie) & "'"
CAD = "select cciudad,ccolonia, cdireccion, cnombrefac cnombre, faccliente,facfecha,total,rfc FROM facventa F, catcliente WHERE cclave = faccliente  and NUMFACTURA =  '" & Trim(Factura) & "' AND SERIE = '" & Trim(SERIE) & "'"
rsFac.Open CAD, cn, adOpenKeyset, adLockOptimistic, 1
'cad = "select cciudad,ccolonia, cdireccion, cnombre, faccliente,facfecha,total,rfc,producto,cantidad,cantidadp,preciop,precio,costo,costop,d.iva, d.ieps, D.IMPORTE, d.tasaieps,d.factura,d.serie,descripc, contenid,medida, paquetes FROM facventa F, facventa_det D, TFPRODUC  T , catcliente WHERE cclave = faccliente  and NUMFACTURA = FACTURA AND F.SERIE = D.SERIE AND D.PRODUCTO = T.CONSEC AND FACTURA = '" & Trim(factura) & "' AND d.SERIE = '" & Trim(serie) & "'"
'rsFac.Open cad, cn, adOpenDynamic, adLockOptimistic, adCmdText
If IsNull(rsFac!rfc) Then
    MsgBox "Es necesario Especificar un rfc para la factura" & vbCrLf & "Por Favor Escriba el rfc Correcto", vbInformation
    RESP = InputBox("Escriba el RFC correcto", "RFC", "COOF970101111")
    rfc = RESP '"COOF970101111"
    CAD = "UPDATE FACVENTA SET RFC =  '" & Trim(rfc) & "'   WHERE NUMFACTURA = '" & Trim(Factura) & "'  AND SERIE = '" & Trim(SERIE) & "'"
    cn.Execute CAD
Else
    rfc = rsFac!rfc
End If

Printer.CurrentY = 3.8 + nAlto
Printer.CurrentX = 0.5
Printer.Print IIf(rfc = "COOF970101111", "C O N S U M I D O R   F I N A L", rsFac!cNombre);
Printer.CurrentX = 10
Printer.Print rfc
Printer.CurrentX = 0.5
Printer.Print IIf(rsFac!rfc = "COOF970101111", "C O N O C I D O", rsFac!cdireccion);
Printer.CurrentX = 10
Printer.Print " " 'rscli!cTelefono
Printer.CurrentX = 0.5
Printer.Print rsFac!ccolonia;
Printer.CurrentX = 10
Printer.Print rsFac!cciudad
Printer.CurrentX = 0.5
Printer.Print Trim(SERIE) & "  " & Factura;
Printer.CurrentY = 6 + nAlto
Printer.CurrentX = 15.5
Printer.Print Format(rsFac!FACFECHA, "long date") & Space(3) & Format(Time, "HH:MM AM/PM")
NVENTA1 = 0: NIVA1 = 0: NIEPS1 = 0
NVENTA2 = 0: NIVA2 = 0: NIEPS2 = 0
NVENTA3 = 0: NIVA3 = 0: NIEPS3 = 0
NVENTA4 = 0: NIVA4 = 0: NIEPS4 = 0
NVENTA5 = 0: NIVA5 = 0: NIEPS5 = 0
NVENTA6 = 0: NIVA6 = 0: NIEPS6 = 0
NVENTA7 = 0: NIVA7 = 0: NIEPS7 = 0
NVENTA8 = 0: NIVA8 = 0: NIEPS8 = 0
Printer.CurrentY = 6.5 + nAlto
rsFac.Close
'EL DETALLE
CAD = "SELECT producto,cantidad,cantidadp,preciop,precio,costo,costop,d.iva, d.ieps, D.IMPORTE, d.tasaieps,d.factura,d.serie,descripc, LTRIM(STR(T.paquetes)) + ' X ' +  lTrim(str(T.contenid,10,3)) + space(2) + SUBSTRING(T.medida,1,5)  AS medida , paquetes FROM facventa_det D, TFPRODUC  T  WHERE D.PRODUCTO = T.CONSEC AND FACTURA = '" & Trim(Factura) & "' AND d.SERIE = '" & Trim(SERIE) & "'"
rsFac.Open CAD, cn, adOpenKeyset, adLockOptimistic, 1
rsFac.MoveFirst

While Not rsFac.EOF
    Printer.CurrentX = 0.3
    If rsFac!cantidad > 0 Then
       Printer.Print rsFac!cantidad & "CJ" & IIf(rsFac!cantidadp > 0, "-" & rsFac!cantidadp & "PZ", "");
    Else
       Printer.CurrentX = 0.3
       Printer.Print rsFac!cantidadp & "PZ";
    End If
    
    Printer.CurrentX = 3
    If Printer.TextWidth(rsFac!descripc) > 5.5 Then
       For N = 1 To Len(rsFac!descripc)
          If Printer.TextWidth(Mid(rsFac!descripc, 1, N)) > 5 Then Exit For
       Next
       Printer.Print Mid(rsFac!descripc, 1, N);
    Else
       Printer.Print rsFac!descripc;
    End If
    
    Printer.CurrentX = 11
    Printer.Print rsFac!medida;
    Printer.CurrentX = 15
    Printer.Print Format(rsFac!ieps, "00");
    Printer.CurrentX = 16
    Printer.Print Format(rsFac!iva, "00");
    Printer.CurrentX = 17
    Printer.Print String(10 - Len(Trim(Format(rsFac!PRECIO, "########0.00"))), " ") & Format(rsFac!PRECIO, "########0.00");
    Printer.CurrentX = 18.5
    Printer.Print String(12 - Len(Trim(Format(rsFac!importe, "########0.00"))), " ") & Format(rsFac!importe, "########0.00")
    ncosto = rsFac!importe
    
    Dim iva  As Currency
    iva = IIf(ZONA = "OAX", 15, 10)
    iva = iva / 100

    If rsFac!tasaieps = 1 Then        'Depto 1
       NVENTA1 = NVENTA1 + ncosto
       NIVA1 = 0
       NIEPS1 = 0
    ElseIf rsFac!tasaieps = 2 Then   'Depto 2
        NVENTA2 = NVENTA2 + ncosto
        NIVA2 = NIVA2 + (ncosto / (1 + iva) * iva)
        NIEPS2 = 0
    ElseIf rsFac!tasaieps = 3 Then  'Depto 3
        NVENTA3 = NVENTA3 + ncosto
        NIVA3 = NIVA3 + (ncosto / (1 + iva) * iva)
        NIEPS3 = NIEPS3 + (((ncosto / (1 + iva)) / 1.25) * 25 / 100)
    ElseIf rsFac!tasaieps = 4 Then 'Depto 4
        NVENTA4 = NVENTA4 + ncosto
        NIVA4 = NIVA4 + (ncosto / (1 + iva) * iva)
        NIEPS4 = NIEPS4 + (((ncosto / (1 + iva)) / 1.3) * 30 / 100)
    ElseIf rsFac!tasaieps = 5 Then 'Depto 5
        NVENTA5 = NVENTA5 + ncosto
        NIVA5 = NIVA5 + (ncosto / (1 + iva) * iva)
        NIEPS5 = 0
    ElseIf rsFac!tasaieps = 6 Then 'Depto 6
        NVENTA6 = NVENTA6 + ncosto
        NIVA6 = NIVA6 + (ncosto / (1 + iva) * iva)
        NIEPS6 = NIEPS6 + (((ncosto / (1 + iva)) / 1.5) * 50 / 100)
    ElseIf rsFac!tasaieps = 7 Then 'Depto 7
        NVENTA7 = NVENTA7 + ncosto
        NIVA7 = NIVA7 + (ncosto / (1 + iva) * iva)
        NIEPS7 = NIEPS7 + (((ncosto / (1 + iva)) / 1.05) * 5 / 100)
    ElseIf rsFac!tasaieps = 8 Then  'Depto 8
        NVENTA8 = NVENTA8 + ncosto
        NIVA8 = NIVA8 + (ncosto / (1 + iva) * iva)
        NIEPS8 = NIEPS8 + (((ncosto / (1 + iva)) / 1.2) * 20 / 100)
    End If
    nProd = nProd + 1
    rsFac.MoveNext
Wend
Printer.CurrentY = 11.5 + nAlto
Printer.CurrentX = 1
Printer.Print "DEP1";
Printer.CurrentX = 3
Printer.Print "DEP2";
Printer.CurrentX = 5
Printer.Print "DEP3";
Printer.CurrentX = 7
Printer.Print "DEP4";
'Printer.CurrentY = 16.5 + nAlto
Printer.CurrentX = 9
Printer.Print "DEP5";
'Printer.CurrentY = 16.5 + nAlto
Printer.CurrentX = 11
Printer.Print "DEP6";
'Printer.CurrentY = 16.5 + nAlto
Printer.CurrentX = 13
Printer.Print "DEP7"
'SUBTOTALES
'Cuando es consumidor final no se desglosa la factura ni se imprime ieps e iva

Printer.CurrentX = 1
Printer.Print Format(NVENTA1, "########0.00");
Printer.CurrentX = 3
'En el caso de que sea consumidor final no se presenta el ieps y iva
Printer.Print Format(NVENTA2, "########0.00");
Printer.CurrentX = 5
Printer.Print Format(NVENTA3, "########0.00");
Printer.CurrentX = 7
Printer.Print Format(NVENTA4, "########0.00");
Printer.CurrentX = 9
Printer.Print Format(NVENTA5, "########0.00");
Printer.CurrentX = 11
Printer.Print Format(NVENTA6, "########0.00");
Printer.CurrentX = 13
Printer.Print Format(NVENTA7, "########0.00");

Printer.CurrentX = 16
Printer.Print "SUBTOTAL";
Printer.CurrentX = 18.5
Printer.FontSize = 10
Printer.Print String(10 - Len(Trim(Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4) + (NVENTA5 - NIVA5 - NIEPS5) + (NVENTA6 - NIVA6 - NIEPS6) + (NVENTA7 - NIVA7 - NIEPS7), "#########0.00"))), " ") & Format(NVENTA1 + (NVENTA2 - NIVA2) + (NVENTA3 - NIVA3 - NIEPS3) + (NVENTA4 - NIVA4 - NIEPS4), "#########0.00")
Printer.FontSize = 8
'IVA
Printer.CurrentX = 1
'CONDICIONES CUANDO NO ES CONSUMIDOR FINAL
Printer.Print Format(NIVA1, "########0.00");
Printer.CurrentX = 3
Printer.Print Format(NIVA2, "########0.00");
Printer.CurrentX = 5
Printer.Print Format(NIVA3, "########0.00");
Printer.CurrentX = 7
Printer.Print Format(NIVA4, "########0.00");
Printer.CurrentX = 9
Printer.Print Format(NIVA5, "########0.00");
Printer.CurrentX = 11
Printer.Print Format(NIVA6, "########0.00");
Printer.CurrentX = 13
Printer.Print Format(NIVA7, "########0.00");
Printer.CurrentX = 16
Printer.Print "IVA";
Printer.CurrentX = 18.5
Printer.FontSize = 10
Printer.Print String(10 - Len(Trim(Format(NIVA2 + NIVA3 + NIVA4 + NIVA5 + NIVA6 + NIVA7, "#########0.00"))), " ") & Format(NIVA2 + NIVA3 + NIVA4, "#########0.00")
Printer.FontSize = 8
'IEPS
Printer.CurrentX = 1
Printer.Print Format(NIEPS1, "########0.00");
Printer.CurrentX = 3
Printer.Print Format(NIEPS2, "########0.00");
Printer.CurrentX = 5
Printer.Print Format(NIEPS3, "########0.00");
Printer.CurrentX = 7
Printer.Print Format(NIEPS4, "########0.00");
Printer.CurrentX = 9
Printer.Print Format(NIEPS5, "########0.00");
Printer.CurrentX = 11
Printer.Print Format(NIEPS6, "########0.00");
Printer.CurrentX = 13
Printer.Print Format(NIEPS7, "########0.00");
Printer.CurrentX = 16
Printer.Print "IEPS";
Printer.CurrentX = 18.5
Printer.FontSize = 10
Printer.Print String(10 - Len(Trim(Format(NIEPS3 + NIEPS4 + NIEPS5 + NIEPS6 + NIEPS7, "#########0.00"))), " ") & Format(NIEPS3 + NIEPS4, "#########0.00")
Printer.FontSize = 8

Printer.FontSize = 8
'TOTAL DE LA VENTA
Printer.CurrentX = 1
If NVENTA1 > 0 Then Printer.Print Format(NVENTA1, "########0.00");
Printer.CurrentX = 3
If NVENTA2 > 0 Then Printer.Print Format(NVENTA2, "########0.00");
Printer.CurrentX = 5
If NVENTA3 > 0 Then Printer.Print Format(NVENTA3, "########0.00");
Printer.CurrentX = 7
If NVENTA4 > 0 Then Printer.Print Format(NVENTA4, "########0.00");
Printer.CurrentX = 9
If NVENTA5 > 0 Then Printer.Print Format(NVENTA5, "########0.00");
Printer.CurrentX = 11
If NVENTA6 > 0 Then Printer.Print Format(NVENTA6, "########0.00");
Printer.CurrentX = 13
If NVENTA7 > 0 Then Printer.Print Format(NVENTA7, "########0.00");

Printer.CurrentX = 16
Printer.Print "TOTAL";
Printer.CurrentX = 18.5
Printer.FontSize = 10
Printer.Print String(10 - Len(Trim(Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7, "#########0.00"))), " ") & Format(NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7, "#########0.00")
Printer.FontSize = 8
Printer.CurrentX = 2
TotFac = NVENTA1 + NVENTA2 + NVENTA3 + NVENTA4 + NVENTA5 + NVENTA6 + NVENTA7
LETRAS = frmVentas.NumLet$(TotFac)
Printer.Print LETRAS
lNvaFac = True: nProd = 0
Printer.EndDoc

MsgBox "LA IMPRESION SE REALIZO CORRECTAMENTE", vbInformation, "Impresión finalizada"
Exit Sub
Error:
  MsgBox Err.Description, vbInformation
End Sub


Private Sub Form_Unload(Cancel As Integer)
 'frmMayoreo.Show
frmAreaRecibo.Show
End Sub


Private Sub mnuact_Click()
Dim RESP, nPos
RESP = InputBox("Proporciona el RFC correcto", "actualiza RFC")
If Trim(RESP) <> "" Or Len(RESP) > 10 Then
   nPos = InStr(1, AdoFacturas.Recordset!Factura, "-")
   cn.Execute "UPDATE facventa SET rfc = '" & UCase(Trim(RESP)) & "' WHERE serie = '" & Trim(Mid(AdoFacturas.Recordset!Factura, 1, nPos - 1)) & "' AND numfactura = '" & Trim(Mid(AdoFacturas.Recordset!Factura, nPos + 1)) & "'"
   cn.Execute "UPDATE facventa_det SET rfc_det = '" & UCase(Trim(RESP)) & "' WHERE serie = '" & Trim(Mid(AdoFacturas.Recordset!Factura, 1, nPos - 1)) & "' AND factura = '" & Trim(Mid(AdoFacturas.Recordset!Factura, nPos + 1)) & "'"
   MsgBox "LA ACTUALIZACION SE REALIZO CORRECTAMENTE", vbInformation, "Facturación"
End If
End Sub

Private Sub mnucobra_Click()
fecha = InputBox("Introduzca fecha a desactivar " & Chr(13) & Chr(13) & "Formato (dd/mm/aaaa)", "Abonos", date)
If IsDate(fecha) Then
   cn.Execute "INSERT INTO diasver(fechaveri, fechacap) VALUES ('" & fecha & "','" & date & "')"
   MsgBox "La fecha " & fecha & " se ha inactivado para captura de abonos", vbInformation, "Abonos"
Else
   MsgBox "El dato introducido no corresponde a una fecha", vbInformation, "Abonos"
End If
End Sub

Private Sub mnufaccan_Click()
Dim nfactura As Double
nfactura = InputBox("Número de factura a insertar", "Factura cancelada")
If Not IsNumeric(nfactura) Then Exit Sub
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT numfactura,serie FROM facventa WHERE numfactura = '" & nfactura & "' and serie = '" & SERIE & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
If rs.BOF And rs.EOF Then
   cn.Execute "INSERT INTO facventa(NOVENTA,faccliente,facfecha,total,iva,ieps,numfactura,serie,cancelado) values (" & _
      12 & ",2,'" & date & "',0,0,0,'" & Trim(nfactura) & "','" & Trim(SERIE) & "',1)"
   cn.Execute "INSERT INTO FACVENTA_DET(producto,cantidad,cantidadp,precio,preciop,costo,costop,importe,iva,ieps,tasaieps,serie,factura,venta,rfc_det,fecha_det) values( " & _
                       "'1008833',1,0,0,0,1,1,0,0,0,1,'" & SERIE & "','" & nfactura & "',12,'CANC999999999','" & date & "')"
   MsgBox "Proceso completado", vbInformation, "Facturas"
End If

End Sub

Private Sub mnuFacPag_Click()
Enca = "PAGOS REALIZADOS DEL " & Me.dtpFecha(0).Value & " AL " & Me.dtpFecha(1).Value
cr1.Connect = cCadConex
cr1.ReportFileName = App.Path & "\abonos.rpt"
cr1.WindowTitle = Enca
cr1.Formulas(0) = "ENCAB = '" & Enca & "'"
If MsgBox("Pagos parciales en base a fechas de pago" & Chr(13) & Chr(13) & "S I  = Pagos parciales basados en fechas de pago" & Chr(13) & "NO = Pagos parciales basado en rango de fechas de facturación ", vbQuestion + vbYesNo, "Pagos") = vbYes Then
   CCONDABO = " ABONOS.fecha >= '" & Format(Me.dtpFecha(0).Value, "yyyy-dd-mm") & "' AND ABONOS.fecha <= '" & Format(Me.dtpFecha(1).Value, "yyyy-dd-mm") & "' "
Else
   CCONDABO = " facventa.facfecha >= '" & Format(Me.dtpFecha(0).Value, "yyyy-dd-mm") & "' AND facventa.facfecha <= '" & Format(Me.dtpFecha(1).Value, "yyyy-dd-mm") & "' "
End If
cr1.SQLQuery = "SELECT FACVENTA.facfecha, FACVENTA.total, FACVENTA.numfactura, FACVENTA.serie, FACVENTA.rfc, " & _
                       "CATCLIENTE.cclave, CATCLIENTE.cnombre, " & _
                       "ABONOS.fecha, ABONOS.importe, ABONOS.tipopag, ABONOS.numero, ABONOS.posfechado, " & _
                       "FOLPOLIZA.folio, FOLPOLIZA.poliza" & Chr(13) & _
               "FROM pitico.dbo.FACVENTA FACVENTA, " & _
                       "PITICO.dbo.CATCLIENTE CATCLIENTE, " & _
                       "PITICO.dbo.ABONOS ABONOS, " & _
                       "pitico.dbo.FOLPOLIZA FOLPOLIZA " & Chr(13) & _
               "WHERE  FACVENTA.faccliente = CATCLIENTE.cclave AND FACVENTA.numfactura = ABONOS.factura AND " & _
                       "FACVENTA.serie = ABONOS.serie AND ABONOS.posfechado IS NULL AND " & _
                       "ABONOS.fecha = FOLPOLIZA.fecha AND " & _
                       "ABONOS.serie = FOLPOLIZA.serie AND " & _
                       "FACVENTA.serie = '" & SERIE & "' AND FOLPOLIZA.poliza = 0 AND " & CCONDABO & Chr(13) & _
               "ORDER BY FACVENTA.facfecha ASC, FACVENTA.numfactura ASC"
'MsgBox cr1.SQLQuery
cr1.Action = 1
End Sub

Private Sub mnuimp_Click()
'POSIBILIDAD DE IMPRIMIR
'On Error GoTo Error:
Dim Factura As String
Dim SERIE As String
    nSepara = InStr(1, AdoFacturas.Recordset!Factura, "-")
    SERIE = Mid(AdoFacturas.Recordset!Factura, 1, nSepara - 1)
    Factura = Trim(Mid(AdoFacturas.Recordset!Factura, nSepara + 1, Len(AdoFacturas.Recordset!Factura)))
If ZONA = "OAX" Then
   Call FACTURAIMPRIMIRv(Factura, SERIE)
Else
   Call FACTURAIMPRIMIRh(Factura, SERIE)
End If
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub mnumes_Click()
cr1.WindowTitle = "Reporte de vales y cheques"
cr1.ReportFileName = App.Path & "\VALCHEQ.RPT"
cr1.Formulas(0) = "FORMSELEC = {FACVENTA.facfecha} >= date(" & Format(dtpFecha(0).Value, "YYYY,MM,DD") & ") and {FACVENTA.facfecha} <= date(" & Format(dtpFecha(1).Value, "YYYY,MM,DD") & ")"
cr1.Action = 1
End Sub

Private Sub mnumod_Click()
On Error GoTo Error:
Dim RESP As String
RESP = InputBox("Introduce la serie a la que se cambiará la factura", "Serie")

cn.Execute "UPDATE VENTAS_DET SET SERIE = '" & RESP & "' WHERE SERIE = '" & Trim(AdoFacturas.Recordset!SERIE) & "' AND FACTURA = '" & Trim(AdoFacturas.Recordset!numfactura) & "'"
cn.Execute "UPDATE ABONOS SET SERIE = '" & RESP & "' WHERE SERIE = '" & Trim(AdoFacturas.Recordset!SERIE) & "' AND FACTURA = '" & Trim(AdoFacturas.Recordset!numfactura) & "'"
cn.Execute "UPDATE FACVENTA_DET SET SERIE = '" & RESP & "' WHERE SERIE = '" & Trim(AdoFacturas.Recordset!SERIE) & "' AND FACTURA = '" & Trim(AdoFacturas.Recordset!numfactura) & "'"
cn.Execute "UPDATE FACVENTA SET SERIE = '" & RESP & "' WHERE SERIE = '" & Trim(AdoFacturas.Recordset!SERIE) & "' AND NUMFACTURA = '" & Trim(AdoFacturas.Recordset!numfactura) & "'"
MsgBox "PROCESO FINALIZADO", vbInformation, "Mensaje"
Exit Sub
Error:
MsgBox Err.Description
End Sub


Private Sub mnurep_Click()
cmdOpcion_Click 3
End Sub

Private Sub mnusal_Click()
Unload Me
frmAreaRecibo.Show
End Sub

Private Sub mnusin_Click()
lprov = "SIN"
cmdOpcion_Click 0
End Sub

Private Sub mnuz_Click()
On Error GoTo Error:
Dim RESP
  RESP = InputBox("Introduzca observaciones de la factura", "Observaciones")
  cn.Execute "UPDATE facventa SET concepto = '" & RESP & "' WHERE  rtrim(serie) +'-'+ numfactura = '" & AdoFacturas.Recordset!Factura & "'"
  MsgBox "La observación se grabo correctamente", vbInformation, "Facturas"
Exit Sub
Error:
   MsgBox Err.Description
End Sub


