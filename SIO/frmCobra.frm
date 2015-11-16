VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCobra 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobrar venta"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmCobra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraVales 
      Caption         =   "Vales"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3135
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ListBox lstVales 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   810
         ItemData        =   "frmCobra.frx":030A
         Left            =   420
         List            =   "frmCobra.frx":030C
         TabIndex        =   18
         Top             =   905
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdcierraval 
         BackColor       =   &H0080FFFF&
         Cancel          =   -1  'True
         Caption         =   "&X"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cierra captura de vales"
         Top             =   0
         Width           =   255
      End
      Begin MSAdodcLib.Adodc AdoVales 
         Height          =   330
         Left            =   360
         Top             =   2640
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
         Caption         =   "AdoVales"
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
      Begin MSDataGridLib.DataGrid dbgrdVales 
         Bindings        =   "frmCobra.frx":030E
         Height          =   1815
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   -2147483624
         BorderStyle     =   0
         ForeColor       =   8388608
         HeadLines       =   1.4
         RowHeight       =   15
         TabAction       =   2
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "noventa"
            Caption         =   "noventa"
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
            DataField       =   "empresa"
            Caption         =   "Empresa"
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
            DataField       =   "cantidad"
            Caption         =   "Cant."
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
            DataField       =   "monto"
            Caption         =   "Importe"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Button          =   -1  'True
               ColumnWidth     =   2129.953
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   524.976
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbgrdCheque 
         Bindings        =   "frmCobra.frx":0325
         Height          =   1815
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   14737632
         BorderStyle     =   0
         ForeColor       =   8388608
         HeadLines       =   1.4
         RowHeight       =   15
         TabAction       =   2
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "noventa"
            Caption         =   "noventa"
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
            DataField       =   "empresa"
            Caption         =   "Banco"
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
            DataField       =   "numcheque"
            Caption         =   "3 digit."
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
            DataField       =   "monto"
            Caption         =   "Importe"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Button          =   -1  'True
               ColumnWidth     =   2129.953
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtcheques 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "cheques"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adoventas"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   870
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   2280
      Width           =   2500
   End
   Begin VB.TextBox txtvales 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "vales"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adoventas"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   870
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   1320
      Width           =   2500
   End
   Begin MSComctlLib.StatusBar STB1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7410
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   661
      SimpleText      =   "               Para salir  presione la tecla  [Esc]"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5539
            MinWidth        =   5539
            Text            =   "Esc => Salir"
            TextSave        =   "Esc => Salir"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "V => Vales"
            TextSave        =   "V => Vales"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "C => Cheques"
            TextSave        =   "C => Cheques"
         EndProperty
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
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cerrar"
      Height          =   465
      Left            =   3240
      Picture         =   "frmCobra.frx":033C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Cobrar"
      Height          =   465
      Left            =   1320
      Picture         =   "frmCobra.frx":04AE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtefectivo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "efectivo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      DataSource      =   "AdoVentas"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   870
      Left            =   3000
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   3240
      Width           =   2500
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   870
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   4320
      Width           =   2500
   End
   Begin VB.TextBox txtsubtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   870
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   120
      Width           =   2500
   End
   Begin VB.TextBox txtCambio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   870
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   5640
      Width           =   2500
   End
   Begin VB.Image imgcobra 
      Height          =   495
      Index           =   6
      Left            =   240
      Picture         =   "frmCobra.frx":0620
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image imgcobra 
      Height          =   495
      Index           =   5
      Left            =   240
      Picture         =   "frmCobra.frx":12C2
      Top             =   360
      Width           =   480
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   120
      X2              =   5520
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   120
      X2              =   5520
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   120
      X2              =   5520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   120
      X2              =   5520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image imgcobra 
      Height          =   495
      Index           =   4
      Left            =   360
      Picture         =   "frmCobra.frx":1F64
      Top             =   4440
      Width           =   510
   End
   Begin VB.Image imgcobra 
      Height          =   480
      Index           =   3
      Left            =   240
      Picture         =   "frmCobra.frx":2D0E
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   600
   End
   Begin VB.Image imgcobra 
      Height          =   480
      Index           =   2
      Left            =   240
      Picture         =   "frmCobra.frx":39B0
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image imgcobra 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmCobra.frx":475A
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label lbletiquetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C&heques"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Index           =   17
      Left            =   1200
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lbletiquetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Vales Despensa"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   15
      Left            =   1200
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lbletiquetas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Efectivo"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   11
      Left            =   1200
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lbletiquetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambio"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   12
      Left            =   1320
      TabIndex        =   8
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label lbletiquetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total a pagar"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   10
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblImporte 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tot. Pago"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "frmCobra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancellar_Click()
  Unload Me
End Sub

Private Sub chkOficial_Click()
  cmbCheques.Visible = (chkOficial.Value = 1)
  If chkOficial.Value = 0 And Me.cmdGrabar.Enabled = True Then
     lbletiquetas(0).Caption = "Ult. digitos"
     lbletiquetas(0).Refresh
     cmdAddCheVal.Enabled = True
     'Me.txtCant.SetFocus
  End If
End Sub

Private Sub cmbCheques_GotFocus()
RESP = SendMessageLong(cmbCheques.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbCheques_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmbCheques_LostFocus()
RESP = SendMessageLong(cmbCheques.hwnd, &H14F, False, 1)
End Sub

Private Sub cmbCheques_Validate(Cancel As Boolean)
 cmbVales.Visible = False
 nPos = InStr(1, "1234567891011", Trim(Mid(cmbCheques.Text, 1, 3)))
 If nPos <> 0 And Trim(cmbCheques.Text) <> "" Then
    cmdAddCheVal.Enabled = (cmdGrabar.Enabled = True)
    lbletiquetas(0).Caption = "Ult. digitos"
    lbletiquetas(0).Refresh
    txtCant.SetFocus
 Else
    cmdAddCheVal.Enabled = False
 End If
End Sub

Private Sub cmbVales_GotFocus()
RESP = SendMessageLong(cmbVales.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbVales_KeyPress(KeyAscii As Integer)
If KeyPress = 27 Then
  Me.txtefectivo.SetFocus
  KeyAscii = 0
  SendKeys vbTab
Unload Me
End If
End Sub

Private Sub cmbVales_LostFocus()
RESP = SendMessageLong(cmbVales.hwnd, &H14F, False, 1)
RESP = SendMessageLong(cmbCheques.hwnd, &H14F, False, 1)
End Sub


Private Sub cmbVales_Validate(Cancel As Boolean)
 nPos = InStr(1, "12345678", Trim(Mid(cmbVales.Text, 1, 3)))
 If nPos <> 0 And Trim(cmbVales.Text) <> "" Then
    If Mid(cmbVales.Text, 1, 1) = "8" Then
        cmdAddCheVal.Enabled = (cmdGrabar.Enabled = True)
        lbletiquetas(0).Caption = "Folio"
        lbletiquetas(0).Refresh
        txtDenom.Text = "190": txtDenom.Enabled = False
        txtCant.SetFocus
    Else
        cmdAddCheVal.Enabled = (cmdGrabar.Enabled = True)
        cmbCheques.Text = ""
        cmbCheques.Visible = False
        lbletiquetas(0).Caption = "Cantidad"
        lbletiquetas(0).Refresh
        txtCant.Text = ""
        txtCant.SetFocus
        'SendKeys vbTab
    End If
  Else
    cmdAddCheVal.Enabled = False
 End If
End Sub

Private Sub cmdAddCheVal_Click()
'On Error GoTo error:
 cn.BeginTrans
 'If Not IsNumeric(txtDenom.Text) Or Val(txtDenom.Text) <= 0 And Me.cmbCheques.Visible Then
 If Not IsNumeric(txtDenom.Text) Or Val(txtDenom.Text) <= 0 Then
    MsgBox "ES NECESARIO ESPECIFICAR LA DENOMINACION DEL VALE / CHEQUE ", vbExclamation
    txtDenom.SetFocus
    cn.CommitTrans
    Exit Sub
 End If
 If Not IsNumeric(txtCant.Text) Or Val(txtCant.Text) <= 0 Then
    MsgBox "ES NECESARIO ESPECIFICAR LA CANTIDAD DEL VALE / CHEQUE ", vbExclamation
    txtCant.SetFocus
    cn.CommitTrans
    Exit Sub
 End If
 'Se valida que el monto de los vales sea menor al monto ya que no se puede dar cambio a vales
     
 If Mid(cmbVales.Text, 1, 1) = "8" Then
    If Val(txtvales.Text) + (Val(txtDenom.Text) * 1) > Val(txtsubtotal.Text) And cmbVales.Visible Then
        MsgBox "LA CANTIDAD EN VALES DEBE SER MENOR AL MONTO DE LA VENTA", vbInformation
        'txtDenom.SetFocus
        cn.CommitTrans
        Exit Sub
    End If
 Else
    If Val(txtvales.Text) + Val(txtDenom.Text) * Val(txtCant.Text) > Val(txtsubtotal.Text) And cmbVales.Visible Then
        MsgBox "LA CANTIDAD EN VALES DEBE SER MENOR AL MONTO DE LA VENTA", vbInformation
        'txtDenom.SetFocus
        cn.CommitTrans
        Exit Sub
    End If
 End If
 If cmbVales.Visible Then
    If Mid(cmbVales.Text, 1, 1) = "8" Then
        txtvales.Text = Format(Val(txtvales.Text) + Val(txtDenom.Text) * Val("1"), "########0.00")
        cn.Execute "INSERT INTO ValCheq (noventa, Monto, cantidad, Empresa, numcheque ) VALUES (" & frmVentas.AdoVentas.Recordset!noventa & "," & txtDenom.Text & ",1,'" & cmbVales.Text & "'," & txtCant.Text & ")"
    Else
        txtvales.Text = Format(Val(txtvales.Text) + Val(txtDenom.Text) * Val(txtCant.Text), "########0.00")
        cn.Execute "INSERT INTO ValCheq (noventa, Monto, cantidad, Empresa ) VALUES (" & frmVentas.AdoVentas.Recordset!noventa & "," & txtDenom.Text & "," & txtCant.Text & ",'" & cmbVales.Text & "')"
    End If
 Else
    txtcheques.Text = Format(Val(txtcheques.Text) + Val(txtDenom.Text), "########0.00")
    cn.Execute "INSERT INTO ValCheq(Noventa, Monto, Empresa, Numcheque,Oficial) VALUES (" & frmVentas.AdoVentas.Recordset!noventa & "," & txtDenom.Text & ",'" & cmbCheques.Text & "'," & txtCant.Text & "," & chkOficial.Value & ")"
 End If
 txtImporte.Text = Format(Val(txtvales.Text) + Val(txtcheques.Text) + Val(txtefectivo.Text), "########0.00")
 frmVentas.AdoVentas.Recordset!VALES = txtvales.Text
 frmVentas.AdoVentas.Recordset!cheques = IIf(IsNumeric(txtcheques.Text), txtcheques.Text, 0)
 frmVentas.AdoVentas.Recordset.Update
 cmdAddCheVal.Enabled = False
 txtDenom.Text = 0
 If cmbCheques.Visible Then
    cmbCheques.SetFocus
 Else
    cmbVales.SetFocus
 End If
 cn.CommitTrans
Exit Sub
Error:
  MsgBox "OCURRIO EL SIGUIENTE ERROR: " & Err.Description & Chr(13) & "A CONTINUACION SE CANCELARAN LOS CAMBIOS REALIZADOS"
  cn.RollbackTrans
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdcierraval_Click()
 Dim rsv As ADODB.Recordset
 Set rsv = New ADODB.Recordset
If Me.dbgrdVales.Visible Then
   rsv.Open "SELECT SUM(total) as Totvales FROM valCheq WHERE numcheque is null And noventa = " & frmVentas.AdoVentas.Recordset!noventa, cn, adOpenDynamic, adLockOptimistic, adCmdText
   If Not (rsv.BOF And rsv.EOF) Then
      txtvales.Text = IIf(IsNull(rsv!TOTVALES), 0, Format(rsv!TOTVALES, "#########0.00"))
      frmVentas.AdoVentas.Recordset!VALES = IIf(IsNull(rsv!TOTVALES), 0#, Format(rsv!TOTVALES, "#########0.00"))
      frmVentas.AdoVentas.Recordset.Update
   End If
Else
   rsv.Open "SELECT SUM(total) as Totcheque FROM valCheq WHERE not numcheque is null AND noventa = " & frmVentas.AdoVentas.Recordset!noventa, cn, adOpenDynamic, adLockOptimistic, adCmdText
   If Not (rsv.BOF And rsv.EOF) Then
      txtcheques.Text = IIf(IsNull(rsv!totcheque), 0#, Format(rsv!totcheque, "#########0.00"))
      frmVentas.AdoVentas.Recordset!cheques = IIf(IsNull(rsv!totcheque), 0#, Format(rsv!totcheque, "#########0.00"))
      frmVentas.AdoVentas.Recordset.Update
   End If
End If
FraVales.Visible = False
lstVales.Visible = False
End Sub

Private Sub cmdGrabar_Click()
Dim rsFac As ADODB.Recordset
Dim RstFranq As ADODB.Recordset
'On Error GoTo Error:
If nOp = 3 Then
   Set RstFranq = New ADODB.Recordset
   RstFranq.Open "SELECT franquicia FROM catcliente WHERE cclave = " & frmVentas.txtcampos(4).Text, cn, adOpenForwardOnly, adLockOptimistic, adCmdText
   If RstFranq!franquicia Then
      RstFranq.Close
      RstFranq.Open "SELECT sum(total) TOTAL, empresa  FROM valcheq WHERE numcheque IS NULL and noventa = " & frmVentas.AdoVentas.Recordset!noventa & " GROUP BY empresa", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
      If Not (RstFranq.BOF And RstFranq.EOF) Then
         NCOM = 0
         While Not RstFranq.EOF
             Select Case Trim(RstFranq!EMPRESA)
                    Case "1  TIENDA PASS"
                         NCOM = NCOM + (1.61 * RstFranq!total / 100)
                    Case "2  EFECTIVALE "
                         NCOM = NCOM + (1.15 * RstFranq!total / 100)
                    Case "3  ACCOR"
                         NCOM = NCOM + (0.58 * RstFranq!total / 100)
                    Case "4  SI VALE"
                         NCOM = NCOM + (1.15 * RstFranq!total / 100)
                    Case "5  DICASA"
                         NCOM = NCOM + (1.15 * RstFranq!total / 100)
             End Select
             RstFranq.MoveNext
         Wend
         MsgBox "AL PAGAR CON VALES SE COBRARA COMISION, EN LA FACTURA APARECERA COMO ABARROTES 2 LA CANTIDAD DE " & Format(Round(NCOM, 2), "$###,##0.00"), vbInformation, "Comisión"
         cn.Execute "DELETE FROM ventas_det WHERE noventa = " & frmVentas.AdoVentas.Recordset!noventa & " and cl_producto = '1008834' "
         cn.Execute "INSERT INTO ventas_det(NoVenta,cl_producto,cantidad,cantidadp,precio,TipoCantidad,Ieps,Iva,precosto,Importe,prebajo,precostop,preciop,TASAIEPS) VALUES (" & frmVentas.AdoVentas.Recordset!noventa & ",'1008834'," & Round(NCOM, 2) & ",0,1,1,0,15,1," & Round(NCOM, 2) * 1 & ",0,1 ,1,2 )"
         frmVentas.AdoDetVta.Refresh
      End If
   End If
   RstFranq.Close
   Set RstFranq = Nothing
End If
cn.BeginTrans
If nOp = 4 Then   'Pago total de la venta a credito
   If Val(txtvales.Text) + Val(txtcheques.Text) + Val(txtefectivo.Text) < Val(txtsubtotal.Text) Then
      MsgBox "EL IMPORTE DE LA VENTA ES MAYOR AL PAGO", vbExclamation
      txtefectivo.SetFocus
      cn.CommitTrans
      Exit Sub
   End If
   txtCambio = Format(Val(Format(txtImporte.Text, "########0.00")) - Val(txtsubtotal.Text), "########0.00")
   cn.Execute "UPDATE facventa SET cobrado = 1 WHERE noventa = " & frmHistCred.AdoHisCre.Recordset!noventa
   cn.Execute "INSERT INTO ABONOS(Noventa,fechapago,Importe) VALUES (" & frmHistCred.AdoHisCre.Recordset!noventa & ",'" & date & " " & Time & "'," & txtImporte.Text & ")"
   cn.Execute "UPDATE ventas SET situacion = 2, Modocredito = 'P',cobro = ' & caja & " ' WHERE noventa = " & frmHistCred.AdoHisCre.Recordset!noventa
ElseIf nOp = 5 Then  'ABono parcial
   If Val(txtvales.Text) + Val(txtcheques.Text) + Val(txtefectivo.Text) < Val(txtsubtotal.Text) Then
      MsgBox "EL IMPORTE DE LA VENTA ES MAYOR AL PAGO", vbExclamation
      txtefectivo.SetFocus
      cn.CommitTrans
      Exit Sub
   End If
Else
   If nOp = 1 Or nOp = 3 Then  'Modificar y cobro de pantalla de ventas
      monto = Val(txtvales.Text) + Val(txtcheques.Text) + Val(txtefectivo.Text)
      If monto < Val(txtsubtotal.Text) And (Not frmVentas.AdoVentas.Recordset!credito And Not frmVentas.AdoVentas.Recordset!Prevta) Then
        MsgBox "EL IMPORTE DE LA VENTA ES MAYOR AL PAGO", vbExclamation, "Ventas"
        txtefectivo.SetFocus
        cn.CommitTrans
        Exit Sub
      End If
   Else  'Desde pago de facturas
      If Val(txtvales.Text) + Val(txtcheques.Text) + Val(txtefectivo.Text) < Val(txtsubtotal.Text) Then
        MsgBox "EL IMPORTE DE LA VENTA ES MAYOR AL PAGO", vbExclamation
        txtefectivo.SetFocus
        cn.CommitTrans
        Exit Sub
      End If
   End If
End If
'nOp = 1   Modificar desde ventas
'nOp = 3   Cobro desde ventas
'nOp = 6   Cobro desde facturas
If nOp = 1 Or nOp = 3 Then
      If frmVentas.AdoVentas.Recordset!credito And nOp <> 6 Then
         frmVentas.AdoVentas.Recordset!cobro = Caja
         frmVentas.AdoVentas.Recordset.Update
      Else
         txtCambio.Text = Format(Val(txtImporte.Text) - Val(txtsubtotal.Text), "########0.00")
         frmVentas.AdoVentas.Recordset!efectivo = txtefectivo.Text
         frmVentas.AdoVentas.Recordset!cheques = txtcheques.Text
         frmVentas.AdoVentas.Recordset!VALES = txtvales.Text
         frmVentas.AdoVentas.Recordset!total = txtImporte.Text
         frmVentas.AdoVentas.Recordset!sobrapago = txtCambio.Text
         frmVentas.AdoVentas.Recordset!cobro = Caja
         frmVentas.AdoVentas.Recordset!facturista = Trim(cUsuario)
         frmVentas.AdoVentas.Recordset.Update
      End If
      cn.CommitTrans
      cmdGrabar.Enabled = False
      Unload frmCobra
      'cn.CommitTrans
      If tipotienda = 4 Then
        frmVentas.cmdTicket_Click   'Imprimir Factura
      ElseIf tipotienda = 3 Then
         CAD = "update ventas set situacion = 2 where noventa = " & frmVentas.AdoVentas.Recordset!noventa
         cn.Execute CAD
      End If
ElseIf nOp = 6 Then
      frmFacturas.AdoFacturas.Recordset!cobrado = 1
      frmFacturas.AdoFacturas.Recordset.Update
      cmdGrabar.Enabled = False
      Unload frmCobra
ElseIf nOp = 5 Then  'Abono parcial
      txtCambio = Format(txtImporte - Val(txtsubtotal.Text), "########0.00")
      cn.Execute "INSERT INTO ABONOS(Noventa,fechapago,Importe) VALUES (" & frmHistCred.AdoHisCre.Recordset!noventa & ",'" & date & " " & Time & "'," & txtImporte.Text & ")"
      cn.Execute "UPDATE ventas SET montopagos = montopagos + " & txtImporte.Text & ", sobrapago = " & txtCambio.Text & ", cobro = '" & Caja & "' WHERE noventa = " & frmHistCred.AdoHisCre.Recordset!noventa
End If
On Error Resume Next
cn.CommitTrans
Exit Sub
Error:
  MsgBox "OCURRIO EL SIGUIENTE ERROR: " & Err.Description & Chr(13) & "A CONTINUACION SE CANCELARAN LOS CAMBIOS REALIZADOS"
  cn.RollbackTrans
End Sub

Private Sub dbgrdCheque_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 3 Then   'cantidad o monto del vale
   AdoVales.Recordset!total = 1 * AdoVales.Recordset!monto
   SendKeys "{DOWN}"
End If
End Sub

Private Sub dbgrdCheque_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex = 1 Then  'Obliga a que seleccionen un producto de la lista desplegable
   Cancel = True
   dbgrdCheque_ButtonClick (ColIndex)
End If
End Sub

Private Sub dbgrdCheque_ButtonClick(ByVal ColIndex As Integer)
verLista dbgrdCheque, ColIndex
End Sub

Private Sub dbgrdCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub dbgrdVales_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 2 Or ColIndex = 3 Then  'cantidad o monto del vale
   AdoVales.Recordset!total = AdoVales.Recordset!cantidad * AdoVales.Recordset!monto
   If ColIndex = 3 Then
      SendKeys "{DOWN}"
   End If
End If
End Sub

Private Sub dbgrdVales_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex = 1 Then  'Obliga a que seleccionen un producto de la lista desplegable
   Cancel = True
   dbgrdVales_ButtonClick (ColIndex)
End If
End Sub

Private Sub dbgrdVales_ButtonClick(ByVal ColIndex As Integer)
 verLista dbgrdVales, ColIndex
End Sub
Private Sub verLista(grid As DataGrid, Index As Integer)
On Error GoTo Error:
Dim L As ListBox
    Select Case Index
       Case 1
            Set L = lstVales
    End Select
    If Index = -1 Then Exit Sub
      With L
          'Abajo (3):
          .Left = grid.Left + grid.Columns(Index).Left
          If Not (AdoVales.Recordset.BOF And AdoVales.Recordset.EOF) Then .Top = grid.Top + grid.RowTop(grid.Row) + grid.RowHeight
          .Width = grid.Columns(Index).Width
          '.ListIndex = 0
          .Visible = True
          .ZOrder 0
          .SetFocus
    End With
   Exit Sub
Error:
    MsgBox Err.Description
End Sub


Private Sub dbgrdVales_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub dbgrdVales_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 lstVales.Visible = False
End Sub

Private Sub Form_Activate()
  txtefectivo.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 86 Then  'V = vales
   FraVales.Visible = True
   FraVales.Caption = "Vales"
   dbgrdVales.Visible = True
   dbgrdCheque.Visible = False
   lstVales.Clear
   lstVales.AddItem "1  TIENDA PASS"
   lstVales.AddItem "2  EFECTIVALE"
   lstVales.AddItem "3  ACCOR"
   lstVales.AddItem "4  SI VALE"
   lstVales.AddItem "5  DICASA"
   lstVales.AddItem "6  PITICO CLIENTES"
   lstVales.AddItem "7  PITICO EMPLEADOS"
   lstVales.AddItem "8  VALES MUNICIPIO"
   'Agrego los vales que se reciben (Estan el la propiedad List de lstvales)
    AdoVales.CommandType = adCmdText
    AdoVales.ConnectionString = cCadConex
    AdoVales.RecordSource = "SELECT * FROM valcheq WHERE numcheque is null AND noventa = " & frmVentas.AdoVentas.Recordset!noventa
    AdoVales.Refresh
    If AdoVales.Recordset.BOF And AdoVales.Recordset.EOF Then dbgrdVales_ButtonClick 1
    dbgrdVales.Enabled = cmdGrabar.Enabled
ElseIf KeyCode = 119 Then
   frmCalc.Show 1
ElseIf KeyCode = 67 Then
    If Not stb1.Panels(3).Enabled Then
       MsgBox "A este cliente no se le permite el pago con cheque" & Chr(13) & "Solicite autorización para aceptar el cheque", vbInformation, "No se acepta cheque"
       'Exit Sub
    End If
    FraVales.Visible = True
    FraVales.Caption = "Cheques"
    dbgrdVales.Visible = False
    dbgrdCheque.Visible = True
    lstVales.Clear
    lstVales.AddItem "1   IEEPO"
    lstVales.AddItem "2   SEP"
    lstVales.AddItem "3   PROCAMPO"
    lstVales.AddItem "4   BANAMEX"
    lstVales.AddItem "5   BANCOMER"
    lstVales.AddItem "6   BBV"
    lstVales.AddItem "7   INVERLAT"
    lstVales.AddItem "8   BITAL"
    lstVales.AddItem "9   SERFIN"
    lstVales.AddItem "10  NACIONAL FINANCIERA"
    lstVales.AddItem "11  BANCRECER"
    lstVales.AddItem "12  OTROS"
    'Agrego los vales que se reciben (Estan el la propiedad List de lstvales)
    AdoVales.CommandType = adCmdText
    AdoVales.ConnectionString = cCadConex
    AdoVales.RecordSource = "SELECT * FROM valcheq WHERE not numcheque is null AND noventa = " & frmVentas.AdoVentas.Recordset!noventa
    AdoVales.Refresh
    dbgrdCheque.Enabled = cmdGrabar.Enabled
ElseIf KeyCode = 69 Then
    txtefectivo.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'KeyAscii = 0
   'SendKeys vbTab
ElseIf KeyAscii = 27 Then
   'Unload Me
End If
End Sub

Private Sub txtCant_GotFocus()
  txtCant.SelLength = 10
End Sub


Private Sub lstVales_DblClick()
    lstVales_KeyPress vbKeyReturn
End Sub

Private Sub lstVales_KeyPress(KeyAscii As Integer)
Dim cveprod As String
Dim N As Integer
    'Habilita el uso de teclado para asignar a la celda o Escape...
    Select Case KeyAscii
        Case vbKeyReturn
             Dim DBG As DataGrid
             Set DBG = IIf(Me.dbgrdVales.Visible, dbgrdVales, dbgrdCheque)
             If AdoVales.Recordset.BOF And AdoVales.Recordset.EOF Then AdoVales.Recordset.AddNew
             DBG.Columns(0).Text = frmVentas.AdoVentas.Recordset!noventa
             DBG.Columns(1).Text = lstVales.Text
             DBG.Columns(2).Text = 0
             DBG.Columns(3).Text = 0
             DBG.Columns(4).Text = 0
             lstVales.Visible = False
             DBG.SetFocus   'Para que se posicione en la columna de cajas
             'SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}"
             keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
        Case vbKeyEscape
             lstVales.Visible = False
    End Select
End Sub

Private Sub txtcheques_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub txtDenom_GotFocus()
  txtDenom.SelLength = 10
End Sub

Private Sub txtefectivo_GotFocus()
  txtefectivo.SelLength = 10
End Sub

Private Sub txtefectivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub txtefectivo_LostFocus()
  txtImporte.Text = Format(Val(txtvales.Text) + Val(txtcheques.Text) + IIf(Trim(txtefectivo.Text) = "", 0, Val(txtefectivo.Text)), "###########0.00")
  txtCambio.Text = Format(txtImporte.Text - Val(txtsubtotal.Text), "########0.00")
End Sub

Private Sub txtvales_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
ElseIf KeyAscii = 27 Then
   Unload Me
End If

End Sub
