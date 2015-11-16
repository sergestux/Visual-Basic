VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPedProv 
   Caption         =   "Pedidos por proveedor pendientes de confimar"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmPedProv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   11880
   Visible         =   0   'False
   WindowState     =   2  'Maximized
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
      Height          =   8175
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   11700
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
         Bindings        =   "frmPedProv.frx":0442
         Height          =   3255
         Left            =   360
         TabIndex        =   56
         Top             =   4560
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5741
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
            Caption         =   "SUCUR."
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
               ColumnWidth     =   764.787
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
      Begin VB.TextBox txtObserva 
         BackColor       =   &H80000000&
         Height          =   495
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   240
         Width           =   8775
      End
      Begin VB.CommandButton cmdHisReg 
         Cancel          =   -1  'True
         Caption         =   "&Regresar"
         Height          =   450
         Left            =   10080
         Picture         =   "frmPedProv.frx":045A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   240
         Width           =   1095
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
         Bindings        =   "frmPedProv.frx":05CC
         Height          =   3615
         Left            =   360
         TabIndex        =   55
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
            Caption         =   "SUCUR."
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
               ColumnWidth     =   780.095
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
   Begin VB.CheckBox chkCampos 
      Caption         =   "&Protección"
      DataField       =   "pp_protect"
      DataSource      =   "AdoPedProve"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   59
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      DataField       =   "pp_montosol"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      DataSource      =   "AdoPedProve"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   8880
      TabIndex        =   48
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      DataField       =   "pp_fecent"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "DD/MM/YYYY hh:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      DataSource      =   "AdoPedProve"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   7080
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtcampos 
      DataField       =   "pp_observa"
      DataSource      =   "AdoPedProve"
      Height          =   615
      Index           =   5
      Left            =   1800
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "pp_fecrecibe"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   0
      Left            =   10320
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "pp_notent"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   3
      Left            =   7080
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "f_factura"
      DataSource      =   "AdoFacturas"
      Height          =   285
      Index           =   2
      Left            =   8760
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRecib 
      Alignment       =   2  'Center
      DataField       =   "f_total"
      DataSource      =   "AdoFacturas"
      Height          =   285
      Index           =   1
      Left            =   10200
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkCampos 
      Caption         =   "Recibido"
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
      Left            =   10440
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbPerCon 
      Height          =   315
      Left            =   2520
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.ComboBox cmbProv 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtcampos 
      DataField       =   "pp_perconfirma"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtcampos 
      DataField       =   "pp_fechagen"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtcampos 
      DataField       =   "pp_proveedor"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      DataField       =   "pp_pedido"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      DataField       =   "pp_fecconfirma"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   4
      Left            =   8400
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   8520
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc AdoPedpro 
      Height          =   330
      Left            =   240
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin MSAdodcLib.Adodc AdoInventario 
      Height          =   330
      Left            =   240
      Top             =   5520
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
   Begin VB.Frame fraPedTie 
      BackColor       =   &H00404000&
      Height          =   3975
      Left            =   1080
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   9855
      Begin VB.PictureBox picBotTie 
         BackColor       =   &H00808080&
         Height          =   3255
         Left            =   8640
         ScaleHeight     =   3195
         ScaleWidth      =   1035
         TabIndex        =   45
         Top             =   480
         Width           =   1100
         Begin VB.CommandButton cmdAjuPed 
            Caption         =   "Aju. &Ofi."
            Height          =   495
            Left            =   120
            Picture         =   "frmPedProv.frx":05E5
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Ajustar pedido de oficinas centrales"
            Top             =   1320
            Width           =   800
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   495
            Left            =   120
            Picture         =   "frmPedProv.frx":06E7
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Cancelar pedido de tienda seleccionado..."
            Top             =   1920
            Width           =   795
         End
         Begin VB.CommandButton cmdRegTie 
            Caption         =   "&C&errar "
            Height          =   495
            Left            =   120
            Picture         =   "frmPedProv.frx":0859
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Cerrar pedidos por tienda"
            Top             =   2520
            Width           =   800
         End
         Begin VB.CommandButton cmdReporte 
            Caption         =   "&Ped. Tie."
            Height          =   495
            Index           =   3
            Left            =   120
            Picture         =   "frmPedProv.frx":09CB
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Vista preeliminar de pedido por tienda"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Aju. &Tie."
            Height          =   495
            Left            =   120
            Picture         =   "frmPedProv.frx":0EFD
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Ajustar pedido de tienda seleccionado..."
            Top             =   720
            Width           =   795
         End
      End
      Begin MSDataGridLib.DataGrid dbgrdPedsol 
         Bindings        =   "frmPedProv.frx":0FFF
         Height          =   3375
         Left            =   360
         TabIndex        =   29
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1.5
         RowHeight       =   15
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
         Caption         =   "PEDIDOS DE TIENDA QUE FORMAN EL PEDIDO POR PROVEEDOR"
         ColumnCount     =   4
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
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "SURTBOD"
            Caption         =   "SURT. BOD."
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
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3225.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   2055.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1019.906
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc AdoDbf 
      Height          =   330
      Left            =   3120
      Top             =   4560
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
   Begin VB.PictureBox crpt 
      Height          =   480
      Left            =   4920
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   60
      Top             =   3120
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   2400
      Top             =   5520
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
   Begin MSAdodcLib.Adodc AdoDetGlo 
      Height          =   330
      Left            =   5040
      Top             =   5520
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
      Left            =   2400
      Top             =   5160
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
      Top             =   5880
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
   Begin MSAdodcLib.Adodc AdoPedSol 
      Height          =   330
      Left            =   2640
      Top             =   5880
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
   Begin VB.PictureBox PicBotones 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11820
      TabIndex        =   14
      Top             =   7500
      Visible         =   0   'False
      Width           =   11880
      Begin VB.CommandButton Command2 
         Caption         =   "&Act."
         Height          =   495
         Left            =   2370
         Picture         =   "frmPedProv.frx":1017
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Actualizar (Obtiene nuevamente los datos del pedido)"
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdCancPedprove 
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   5760
         Picture         =   "frmPedProv.frx":1119
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Cancelar pedido de tienda seleccionado..."
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdActPreTie 
         Caption         =   "Act. $ tie."
         Height          =   495
         Left            =   3960
         Picture         =   "frmPedProv.frx":128B
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Actualizar costo de pedidos por tienda que forman el pedido por proveedor"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Det."
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "frmPedProv.frx":138D
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Reporte detallado de productos solicitados por tienda"
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   6720
         Picture         =   "frmPedProv.frx":18BF
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Grabar datos"
         Top             =   120
         Width           =   800
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Cerrar"
         Height          =   495
         Left            =   7560
         Picture         =   "frmPedProv.frx":1A31
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Regresar a la pantalla de pedidos"
         Top             =   120
         Width           =   800
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Glob."
         Height          =   495
         Index           =   1
         Left            =   1560
         Picture         =   "frmPedProv.frx":1BA3
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Reporte de pedido global"
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdCondCom 
         Caption         =   "&Cond."
         Height          =   495
         Left            =   3120
         Picture         =   "frmPedProv.frx":20D5
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Ver cargos, descuentos y promociones de producto"
         Top             =   120
         Width           =   800
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Mixto"
         Height          =   495
         Index           =   2
         Left            =   840
         Picture         =   "frmPedProv.frx":21D7
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Reporte de pedidos a surtirse en bodega"
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdVerPedTie 
         Caption         =   "Ver Ped."
         Height          =   495
         Left            =   4920
         Picture         =   "frmPedProv.frx":2709
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Ver los pedidos de tienda que forman el pedido por proveedor"
         Top             =   120
         Width           =   800
      End
      Begin VB.Label lblCajPza 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CAJAS: XXX            PIEZAS: XX"
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
         Left            =   8400
         TabIndex        =   44
         Top             =   240
         Width           =   3375
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdPedpro 
      Bindings        =   "frmPedProv.frx":2803
      Height          =   4815
      Left            =   120
      TabIndex        =   26
      Top             =   2400
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1.5
      RowHeight       =   15
      TabAction       =   2
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StbMensajes 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   52
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   14887
            MinWidth        =   14887
            Text            =   "Para salir presione la tecla [ Esc ]"
            TextSave        =   "Para salir presione la tecla [ Esc ]"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   6244
            MinWidth        =   6244
            Text            =   "F4 => Consultar historial del producto"
            TextSave        =   "F4 => Consultar historial del producto"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cmdlg 
      Left            =   0
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      DataField       =   "pp_perprotec"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   8
      Left            =   10680
      TabIndex        =   50
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtcampos 
      DataField       =   "pp_sucursal"
      DataSource      =   "AdoPedProve"
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox cmbtienda 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Sucursal"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   58
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Protección por"
      Height          =   255
      Index           =   7
      Left            =   10440
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbletiquetas 
      Alignment       =   2  'Center
      Caption         =   "Monto"
      Height          =   255
      Index           =   6
      Left            =   8880
      TabIndex        =   47
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Fecha de ent. aprox."
      Height          =   255
      Index           =   5
      Left            =   7080
      TabIndex        =   46
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblDespro 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   27
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Nota de entrada"
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Monto"
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Num. Factura"
      Height          =   255
      Index           =   2
      Left            =   8520
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Persona que confirma"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Fecha de elaboracion"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbletiquetas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave del pedido"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmPedProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lMove As Boolean
Private rstDesPro As ADODB.Recordset
Private Tabla

Private Sub AdoEntPend_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
txtobserva = IIf(IsNull(AdoEntPend.Recordset!pp_observa), "", AdoEntPend.Recordset!pp_observa)
End Sub

Private Sub AdoEntSur_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
'txtobserva = IIf(IsNull(AdoEntSur.Recordset!pp_observa), "", AdoEntSur.Recordset!pp_observa)
End Sub

Private Sub AdoPedpro_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Not lMove Then Exit Sub
If AdoPedpro.Recordset.BOF And AdoPedpro.Recordset.EOF Then Exit Sub
rstDesPro.MoveFirst
rstDesPro.Find "CONSEC = '" & AdoPedpro.Recordset!clave & "'"
If rstDesPro.EOF Then
   lblDespro.Caption = ""
Else
   lblDespro.Caption = Trim(AdoPedpro.Recordset!descripcion) & Chr(13) & " " & Trim(AdoPedpro.Recordset!Present) & "  " & "PROMOCION: " & CStr(rstDesPro!cajas) & "/" & CStr(rstDesPro!encajas)
End If
End Sub

Private Sub cmbtienda_GotFocus()
  RESP = SendMessageLong(cmbtienda.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbtienda_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{Tab}"
End If
End Sub

Private Sub cmbtienda_LostFocus()
On Error GoTo Error:
Dim N As Integer
If nOp = 1 Then
   If cmbtienda.Text = "" Or IsNull(cmbtienda.Text) Then
       MsgBox "Debe seleccionar una sucursal de la lista desplegable", vbExclamation
       cmbtienda.SetFocus
       Cancel = True
   Else
        Dim RST As ADODB.Recordset
        Set RST = New ADODB.Recordset
        RST.Open "SELECT * FROM cattienda WHERE tidescrip = '" & Trim(cmbtienda.Text) & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If RST.EOF Or RST.BOF Then
            MsgBox "Debe seleccionar una tienda de la lista desplegable", vbExclamation
            cmbtienda.SetFocus
            Cancel = True
        Else
            txtcampos(9).Text = RST!ticlave
            txtcampos(9).SetFocus
        End If
        RST.Close
        Set RST = Nothing
   End If
End If
RESP = SendMessageLong(cmbtienda.hwnd, &H14F, False, 1)
Exit Sub
Error:
End Sub

Private Sub cmdHisReg_Click()
  Me.FraHist.Visible = False
  PicBotones.Visible = True
End Sub

Private Sub Command2_Click()
AdoPedpro.Refresh
dbgrdPedpro.Columns(1).Width = 5560
End Sub

Private Sub chkCampos_Click(Index As Integer)
On Error GoTo Error:
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
      End If
      For N = 0 To 3
          If N > 0 Then lblRec(N).Visible = chkCampos(Index).Value = 1
          If N < 3 Then txtRecib(N).Visible = chkCampos(Index).Value = 1
      Next
   End If
End If
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub chkSurtBod_Click()
cn.Execute "UPDATE PEDIDOS SET P_surtbodega = " & chkSurtBod.Value & " WHERE P_PEDIDO =' " & Trim(AdoPedSol.Recordset!Folio) & "'"
AdoPedSol.Refresh
End Sub


Private Sub Cmbprov_GotFocus()
RESP = SendMessageLong(cmbProv.hwnd, &H14F, True, 1)
End Sub

Private Sub cmbProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{Tab}"
End If

End Sub

Private Sub Cmbprov_LostFocus()
RESP = SendMessageLong(cmbProv.hwnd, &H14F, False, 1)
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
            txtcampos(1).SetFocus
        End If
    End If
End If
Exit Sub
Error:
End Sub

Private Sub cmdActPreTie_Click()
 If MsgBox("REALMENTE DESEAS ACTUALIZAR EL COSTO DE LOS PEDIDOS POR TIENDA", vbYesNo + vbInformation) = vbYes Then
    cn.Execute "UPDATE DETALLEFACTURA SET DF_COSTO = COSTO FROM PEDIDOS, DESCPROD WHERE DF_PEDIDO = P_PEDIDO AND P_PEDPROVEEDOR = '" & Trim(txtcampos(0).Text) & "' AND DF_PROD = PRODUCTO"
 End If
End Sub

Private Sub cmdAjuPed_Click()
On Error GoTo Error:
  nOp = 1 'Por defaul pongo nuevo pedido
  frmCaptPed.Caption = "Modificar pedido"
  If Not (AdoPedSol.Recordset.BOF And AdoPedSol.Recordset.EOF) Then
    AdoPedSol.Recordset.MoveFirst
    While Not AdoPedSol.Recordset.EOF
     If Mid(AdoPedSol.Recordset!Folio, 1, 3) = "OFI" Then
        frmCaptPed.Caption = "Capturar nuevo pedido"
        nOp = 0
        SendKeys AdoPedSol.Recordset!Folio
        SendKeys vbTab
        AdoPedSol.Recordset.MoveLast
      End If
      AdoPedSol.Recordset.MoveNext
    Wend
  End If
  cModo = "CAPTURARPEDIDO"
  DEDONDE = "PROVEEDOR"
  frmCaptPed.Show
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdCancelar_Click()
Dim cmen
'On Error Resume Next
  If MsgBox("REALMENTE DESEAS CANCELAR EL PEDIDO CON FOLIO: " & AdoPedSol.Recordset!Folio, _
            vbInformation + vbYesNo) = vbYes Then
     cmen = StbMensajes.SimpleText
     cn.Execute "UPDATE pedidos set p_cancelado = 1 WHERE P_PEDIDO = '" & AdoPedSol.Recordset!Folio & "'"
     'StbMensajes.SimpleText = Space(55) & "Espere un momento actualizando total de cajas y piezas.........."
     'StbMensajes.Refresh
     AdoPedSol.Refresh
    ' AdoPedpro.Refresh
    ' frmPedProv.AdoPedpro.Recordset.MoveFirst
    ' nCaj = 0: nPza = 0: npro = 0
    ' While Not AdoPedpro.Recordset.EOF
    '   nCaj = nCaj + AdoPedpro.Recordset!SolxCaja
    '   nPza = nPza + AdoPedpro.Recordset!SolxPza
    '   npro = npro + 1
    '   AdoPedpro.Recordset.MoveNext
    ' Wend
    ' lblCajPza.Caption = "PRODUCTOS: " & CStr(npro) & Space(5) & "CAJAS : " & CStr(nCaj) & Space(5) & "PIEZAS: " & CStr(nPza)
    ' lblCajPza.Refresh
    ' StbMensajes.SimpleText = cMen
    ' StbMensajes.Refresh
  MsgBox "PARA HACER EL RECALCULO DE CAJAS Y PIEZAS SOLICITADAS ES NECESARIO SALIR Y VOLVER A ENTRAR AL PEDIDO", vbInformation
  End If
End Sub

Private Sub cmdCancPedprove_Click()
If MsgBox("CONFIRMA SI DESEAS CANCELAR EL PEDIDO POR PROVEEDOR Y TODOS" & Chr(13) & "LOS PEDIDOS DE TIENDAS QUE LO FORMAN", vbYesNo + vbInformation) = vbYes Then
   cn.Execute "UPDATE pedidos set p_cancelado = 1 WHERE p_pedproveedor  = '" & txtcampos(0).Text & "'"
   cn.Execute "UPDATE pedprove SET pp_cancelado = 1 WHERE pp_pedido = '" & txtcampos(0).Text & "'"
   Unload Me
End If
End Sub

Private Sub cmdCondCom_Click()
  cmen = StbMensajes.SimpleText
  StbMensajes.SimpleText = Space(55) & "Espere un momento cargando cargos, descuentos y promociones"
  StbMensajes.Refresh
  'frmConvenio.txtpedido.Text = txtCampos(0).Text
  strcveprov = txtcampos(0).Text
  'MsgBox txtcampos(0).Text
  'MsgBox dbgrdPedpro.Columns(0).Text
  strcveprod = dbgrdPedpro.Columns(0).Text
  frmConvenio.Show
  StbMensajes.SimpleText = cmen
  StbMensajes.Refresh
End Sub

'Exporta pedidos por proveedor a tablas Dbf de Visual Fox Pro para enviar a Carbonera
'de aquellos productos que seran entregados exclusivamente en Bodega
Private Sub CmdExporta_Click()
On Error GoTo Error:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cnFoxPro As ADODB.Connection
   cMenAnt = StbMensajes.SimpleText
   Cmdlg.DialogTitle = "Grabar archivo para enviar pedidos por proveedor a carbonera"
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
        'Set F = fs.GetFile("C:\PASO\ESTPEDOF.DBF")
        Set f = fs.GetFile("\\SERVIDOR_OAXACA\PROGRAMAS\ESTPEDOF.DBF")
        f.Copy cRutArc, True
   End If

   Set rsttemp = New ADODB.Recordset
   AdoDbf.CursorType = adOpenKeyset
   AdoDbf.CommandType = adCmdText
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch & " ORDER BY Proveed"
   AdoDbf.Refresh
     
   rsttemp.Open "SELECT DF_PROD, SUM(DF_CANTSOL) AS SOLCAJA, SUM(DF_CANTSOLP) AS SOLPZA, MAX(DF_COSTO) AS COSTO FROM PEDIDOS, DETALLEFACTURA, TFPRODUC WHERE P_PEDPROVEEDOR = '" & txtcampos(0).Text & "' AND P_SURTBODEGA = 1 AND P_CANCELADO = 0 AND P_PEDIDO = DF_PEDIDO AND DF_PROD = CONSEC GROUP BY DF_PROD", cn, adOpenStatic, adLockOptimistic, adCmdText
   'MsgBox "A CONTINUACION SE EXPORTARAN: " & CStr(rsttemp.RecordCount) & " REGISTROS", vbInformation
   While Not rsttemp.EOF
      StbMensajes.SimpleText = Space(75) & "Exportando producto con la clave: " & CStr(rsttemp!df_prod)
      StbMensajes.Refresh
      AdoDbf.Recordset.AddNew
      AdoDbf.Recordset!Pedido = txtcampos(0).Text
      AdoDbf.Recordset!producto = rsttemp!df_prod
      AdoDbf.Recordset!CantSolc = rsttemp!solcaja
      AdoDbf.Recordset!cantsolp = rsttemp!solpza
      AdoDbf.Recordset!costo = rsttemp!costo
      AdoDbf.Recordset!proveed = Trim(txtcampos(1).Text)
      AdoDbf.Recordset!FecPed = Trim(txtcampos(3).Text)
      AdoDbf.Recordset!FecConf = Trim(txtcampos(4).Text)
      AdoDbf.Recordset!sucursal = Trim(Mid(cSucursal, 1, 3))
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

Private Sub cmdGrabar_Click()
Dim FecConf As Date
Dim lback As Boolean
Dim rsttemp As ADODB.Recordset
Dim rstPromprod As ADODB.Recordset
Dim lTrans  As Boolean
'On Error GoTo Error
  lMove = False: lback = True ' Flag de Backorder
  'Cuando se confirma
  If chkCampos(0).Value = 0 And nOp = 1 Then
     MsgBox "ES NECESARIO ACTIVAR LA CASILLA DE PEDIDO CONFIRMADO", vbExclamation
     chkCampos(0).SetFocus
     Exit Sub
  End If
  'si se confirma el pedido
  If chkCampos(0).Value = 1 Or chkCampos(0).Value = 0 Then
     PROT = 1
     If IsNumeric(txtcampos(8).Text) Then
        If Val(txtcampos(8).Text) > 0 Then
           MsgBox "ESTE PEDIDO SERA CONSIDERADO DE PROTECCION, Y SE SURTIRAN A LAS TIENDAS POR " & txtcampos(8).Text & " PERIODOS"
           MsgBox "LAS PIEZAS QUE SE SOLICITEN NO SE TOMARAN EN CUENTA SOLAMENTE CAJAS"
           PROT = Val(txtcampos(8).Text)
        End If
     End If
     cn.BeginTrans: lTrans = True
     Set rsttemp = New ADODB.Recordset
     AdoPedProve.Recordset.Update
     AdoPedProve.Recordset!pp_recibe = 0
     AdoPedProve.Recordset.Update
     
     cn.Execute "DELETE FROM DETALLEGLOBAL WHERE DG_PEDIDO = '" & txtcampos(0).Text & "'"
     'Agrego el detalle del pedido por proveedor para prepararlo para recibirlo
     rsttemp.Open "DETALLEGLOBAL", cn, adOpenKeyset, adLockOptimistic, adCmdTable
     lMove = False
     AdoPedpro.Recordset.MoveFirst
     Set rstDesPro = New ADODB.Recordset
     While Not AdoPedpro.Recordset.EOF
         rsttemp.AddNew
         rsttemp!dg_pedido = txtcampos(0).Text
         rsttemp!DG_PRODUCTO = AdoPedpro.Recordset!clave
         rsttemp!dg_cantreal = 0
         rsttemp!dg_cantrealp = 0
         'SE GRABA SOLO LA PARTE QUE DEBE IR A LA BODEGA
         If Not AdoPedProve.Recordset!pp_pedind Then
            nPos = InStr(1, AdoPedpro.Recordset!Present, "X")
            nPaquetes = Val(Mid(AdoPedpro.Recordset!Present, 1, nPos))
            If AdoPedpro.Recordset!SolxPza >= nPaquetes Then
                nCaja = Int(AdoPedpro.Recordset!SolxPza / nPaquetes)
                rsttemp!dg_cantsol = AdoPedpro.Recordset!SolxCaja + nCaja
                rsttemp!dg_cantsolp = (AdoPedpro.Recordset!SolxPza - (nCaja * nPaquetes))
            Else
                If Not IsNull(AdoPedpro.Recordset!SolxCaja) Then rsttemp!dg_cantsol = AdoPedpro.Recordset!SolxCaja * PROT
                If Not IsNull(AdoPedpro.Recordset!SolxPza) Then rsttemp!dg_cantsolp = AdoPedpro.Recordset!SolxPza * PROT
            End If
         Else
            rsttemp!dg_cantsol = AdoPedpro.Recordset!SolxCaj * PROT
            rsttemp!dg_cantsolp = AdoPedpro.Recordset!SolxPza * PROT
         End If
         'Calcula el costo sin promocion del producto
         'nprecio = CostSinProm(AdoPedpro.Recordset!Clave)
         'rsttemp!dg_costo = nprecio
         'rsttemp.Update
         'FALTA AGREGAR LA CONDICION DE PRODUCTO POR PRODUCTO
         CADENA = "SELECT * FROM DescProd WHERE producto = '" & Trim(AdoPedpro.Recordset!clave) & "'"
         rstDesPro.Open CADENA, cn, adOpenKeyset, adLockOptimistic, adCmdText
         If rstDesPro.RecordCount > 0 Then
            'se graban tambien los descuentos y cargo
            rsttemp!dg_costo = rstDesPro!costo
            rsttemp!dg_decto1 = rstDesPro!decto1
            rsttemp!dg_decto2 = rstDesPro!decto2
            rsttemp!dg_decto3 = rstDesPro!decto3
            rsttemp!dg_decto4 = rstDesPro!decto4
            rsttemp!dg_decto5 = rstDesPro!decto5
            rsttemp!dg_decto6 = rstDesPro!financiero
            rsttemp!dg_cargo1 = rstDesPro!cargo1
            rsttemp!dg_cargo2 = rstDesPro!cargo2
            rsttemp!dg_iva = rstDesPro!cargo3
            rsttemp!dg_ieps = rstDesPro!cargo4
            rsttemp!dg_cargo5 = rstDesPro!cargo5
            rsttemp!dg_maniobras = rstDesPro!maniobras
            rsttemp!dg_flete = rstDesPro!flete
            rsttemp!dg_efectivo = rstDesPro!efectivo
            rsttemp!dg_prelista = rstDesPro!preciolista
            rsttemp!dg_cajas = rstDesPro!cajas
            rsttemp!dg_encajas = rstDesPro!encajas
            If rstDesPro!encajas > 0 Then
               rsttemp!dg_promocion = Int(rsttemp!dg_cantsol / rstDesPro!encajas) * rstDesPro!cajas
            End If
         End If
         rsttemp.Update
         rstDesPro.Close
         AdoPedpro.Recordset.MoveNext
     Wend
     If AdoPedProve.Recordset!pp_pedind Then
        cmdReporte(1).Enabled = True
        cmdCondCom.Enabled = True
        cmdActPreTie.Enabled = False
        cmdGrabar.Enabled = True
     Else
         'Confirmo los pedidos por tienda que incluye el Pedido por proveedor
        Set rsttemp = New ADODB.Recordset
        If lNvoPedprove = True Then   'Nuevo pedido
           CADENA = "SELECT * FROM PEDIDOS WHERE P_proveedor = '" & Trim(txtcampos(1).Text) & "'  AND p_situacion = 0 AND p_cancelado = 0"
        Else
           CADENA = "SELECT * FROM PEDIDOS WHERE P_pedproveedor = '" & Trim(txtcampos(0).Text) & "' AND p_cancelado = 0"
           dbgrdPedpro.AllowUpdate = True
        End If
        
        rsttemp.Open CADENA, cn, adOpenKeyset, adLockOptimistic, adCmdText
        While Not rsttemp.EOF
            rsttemp!p_fecConfirma = txtcampos(4).Text
            rsttemp!p_situacion = 1
            rsttemp!P_pedproveedor = txtcampos(0).Text
            rsttemp!p_perconfirma = txtcampos(2).Text
            rsttemp!p_observaciones = Chr(13) & txtcampos(5).Text
            rsttemp.Update
            'TAMBIEN SE DEBEN GRABAR LOS DESCUENTOS Y CARGOS EN CADA PEDIDO SUGERIDO DE TIENDA
            CADENA = "  UPDATE detallefactura SET  dg_decto1  =  descprod.decto1, dg_decto2 =  descprod.decto2 " & _
                ",dg_decto3 =  descprod.decto3 ,dg_decto4 =  descprod.decto4 ,dg_decto5 =  descprod.decto5  " & _
                ",dg_decto6 = descprod.financiero ,dg_cargo1 = descprod.cargo1 ,dg_cargo2 = descprod.cargo2 " & _
                ",dg_iva =   descprod.cargo3 ,dg_ieps = descprod.cargo4 ,dg_cargo5 =  descprod.cargo5 " & _
                ",dg_maniobras = descprod.maniobras ,dg_flete = descprod.flete ,dg_efectivo =  descprod.efectivo " & _
                ",df_costo = descprod.costo,dg_prelista =  descprod.preciolista ,dg_cajas = descprod.Cajas ,dg_encajas =  descprod.Encajas " & _
                "  from descprod where producto = df_prod and df_pedido  = '" & Trim(rsttemp!p_Pedido) & "'"
             cn.Execute CADENA
             rsttemp.MoveNext
          Wend
        rsttemp.Close
        rsttemp.Open "SELECT SUM(dg_cantsol * dg_costo + (dg_costo / paquetes * dg_cantsolp)) AS MonSol FROM detalleglobal,tfproduc WHERE dg_pedido = '" & Trim(txtcampos(0).Text) & "' AND dg_producto = consec", cn, adOpenKeyset, adLockOptimistic, adCmdText
        cn.Execute "UPDATE pedprove SET pp_montosol = " & IIf(IsNull(rsttemp!Monsol), 0, rsttemp!Monsol) & " WHERE pp_pedido = '" & Trim(txtcampos(0).Text) & "'"
        txtcampos(7).Text = Format(IIf(IsNull(rsttemp!Monsol), 0, rsttemp!Monsol), "$ ##,###,###,##0.00")
 
        AdoPedpro.Recordset.MoveFirst
        cmdGrabar.Enabled = True
        cmdCancPedprove.Enabled = True
        cmdReporte(0).Enabled = True
        cmdReporte(1).Enabled = True
        cmdActPreTie.Enabled = True
        If lTrans Then cn.CommitTrans
        lTrans = False
        'Cargo pedidos ya grabados para que puedan modificar el lugar de entrega
        AdoPedSol.CommandType = adCmdText
        AdoPedSol.ConnectionString = cCadConex
        AdoPedSol.RecordSource = "SELECT p_pedido as FOLIO, tidescrip AS SUCURSAL, P_fecPed As FECHA_SOL, P_surtbodega as SURTBOD FROM PEDIDOS,CATTIENDA WHERE P_SUCURSAL = TICLAVE AND P_PEDPROVEEDOR = '" & txtcampos(0).Text & "' ORDER BY SUCURSAL"
        AdoPedSol.Refresh
        'Se cargan los productos del pedido de carbonera exclusivamente
        cCadena = "SELECT dg_producto AS CLAVE, descripc As DESCRIPCION, str(paquetes) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT, InCant AS EXISTENCIA, dg_cantsol AS SOLxCAJA, dg_cantsolp AS SOLxPZA , incant as existencia FROM detalleglobal,tfproduc,inventario WHERE dg_producto = consec and dg_producto =  inprod AND dg_pedido = '" & txtcampos(0).Text & "' ORDER BY descripc,contenid"
        AdoPedpro.ConnectionString = cCadConex
        AdoPedpro.CommandType = adCmdText
        AdoPedpro.RecordSource = cCadena
        AdoPedpro.Refresh
        dbgrdPedpro.Columns(1).Width = 5560
        dbgrdPedpro.Columns(0).Locked = True
        dbgrdPedpro.Columns(1).Locked = True
        dbgrdPedpro.Columns(2).Locked = True
        dbgrdPedpro.AllowUpdate = True
        lNvoPedprove = False    'Para que puedan modificar la cantidad solicitada en piezas global
     End If
     rsttemp.Close
     rsttemp.Open "SELECT COUNT(DG_PRODUCTO) AS TOTPRO, SUM(DG_CANTSOL) AS TOTCAJ, SUM(DG_CANTSOLP) AS TOTPZA FROM DETALLEGLOBAL WHERE dg_pedido = '" & txtcampos(0).Text & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
     lblCajPza.Visible = True
     lblCajPza.Caption = "PRODUCTOS: " & CStr(rsttemp!totpro) & Space(5) & "CAJAS:  " + CStr(rsttemp!totcaj) + Space(5) + "PIEZAS:  " + CStr(rsttemp!totpza)
     cmdGrabar.Enabled = True 'TEMPSAVE Mientras se acostumbran a que un avez grabado ya no podran modificar
     chkCampos(0).Enabled = False
     If lTrans Then cn.CommitTrans
     AdoPedProve.RecordSource = "SELECT * FROM Pedprove WHERE pp_pedido = '" & txtcampos(0).Text & "'"
     AdoPedProve.Refresh   '
     lMove = True  'Activo nuevamente la etiqueta de promociones
  End If
Exit Sub
Error:
   If lTrans Then
      MsgBox "OCURRIO EL SIGUIENTE ERROR: " & Err.Description, vbInformation
      MsgBox "A CONTINUACION SE DESAHARAN LOS CAMBIOS REALIZADOS", vbCritical
      cn.RollbackTrans
   End If
   MsgBox Err.Description
End Sub

Private Sub cmdRegresar_Click()
  Unload Me
End Sub

Private Sub cmdRegTie_Click()
   dbgrdPedpro.Visible = True
   PicBotones.Visible = True
   fraPedTie.Visible = False
End Sub

Private Sub cmdReporte_Click(Index As Integer)
Dim rsttemp As ADODB.Recordset
On Error GoTo Error:
cMensaje = StbMensajes.SimpleText
StbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
StbMensajes.Refresh
crpt.Connect = cCadConex
 Select Case Index
 Case 0 'Detallado
      cARcRpt = App.Path & "\Pedprove.rpt"
      'ccondrpt = "FORMSELEC = {PEDPROVE.pp_pedido} = '" & txtcampos(0).Text & "'"
      crpt.SQLQuery = "SELECT PEDPROVE.pp_pedido, PEDPROVE.pp_observa, " & _
                             "PEDIDOS.p_pedido, PEDIDOS.p_sucursal, " & _
                             "DETALLEFACTURA.df_prod, DETALLEFACTURA.df_cantsol, DETALLEFACTURA.df_cantsolp, " & _
                             "TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES " & Chr(13) & _
                      "FROM pitico.dbo.PEDPROVE PEDPROVE, " & _
                            "pitico.dbo.PEDIDOS PEDIDOS, " & _
                            "pitico.dbo.DETALLEFACTURA DETALLEFACTURA, " & _
                            "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                      "WHERE PEDPROVE.pp_pedido = PEDIDOS.p_pedproveedor AND " & _
                            "PEDIDOS.p_pedido = DETALLEFACTURA.df_pedido AND " & _
                            "DETALLEFACTURA.df_prod = TFPRODUC.CONSEC AND " & _
                            "PEDPROVE.pp_pedido = '" & Trim(txtcampos(0).Text) & "' " & Chr(13) & _
                      "ORDER BY DETALLEFACTURA.df_prod ASC "
      crpt.Formulas(5) = "PROVED = '" & Me.cmbProv.Text & "'"
 Case 1 'Global
     If Mid(txtcampos(0).Text, 1, 3) = "C33" Or Mid(txtcampos(0).Text, 1, 3) = "C37" Then
        cARcRpt = App.Path & "\Pedprovg2.rpt"
     Else
        cARcRpt = App.Path & "\Pedprovg1.rpt"
     End If
     ccondrpt = "FORMSELEC = {PEDPROVE.pp_pedido} = '" & txtcampos(0).Text & "'"
     crpt.SQLQuery = "SELECT PEDPROVE.pp_proveedor, PEDPROVE.pp_pedido, PEDPROVE.pp_observa, " & _
                             "DETALLEGLOBAL.dg_producto, DETALLEGLOBAL.dg_cantsol, DETALLEGLOBAL.dg_cantsolp, DETALLEGLOBAL.dg_costo, DETALLEGLOBAL.dg_decto1, DETALLEGLOBAL.dg_decto2, DETALLEGLOBAL.dg_decto3, DETALLEGLOBAL.dg_decto4, DETALLEGLOBAL.dg_decto5, DETALLEGLOBAL.dg_decto6, DETALLEGLOBAL.dg_iva, DETALLEGLOBAL.dg_ieps, DETALLEGLOBAL.dg_cajas, DETALLEGLOBAL.dg_encajas, DETALLEGLOBAL.dg_prelista, DETALLEGLOBAL.dg_efectivo, " & _
                             "CATPROV.NOMPROVE, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.clavedelprov " & Chr(13) & _
                      "From pitico.dbo.PEDPROVE PEDPROVE, " & _
                            "pitico.dbo.DETALLEGLOBAL DETALLEGLOBAL, " & _
                            "PITICO.dbo.USUARIOS USUARIOS, " & _
                            "pitico.dbo.CATPROV CATPROV, pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                      "Where PEDPROVE.pp_pedido = DETALLEGLOBAL.dg_pedido AND " & _
                             "PEDPROVE.pp_proveedor = CATPROV.PROVE AND " & _
                             "DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND " & _
                             "CATPROV.comprador = USUARIOS.LOGIN AND " & _
                             "(DETALLEGLOBAL.dg_cantsol > 0 OR DETALLEGLOBAL.dg_cantsolP > 0 ) AND " & _
                             "PEDPROVE.pp_pedido = '" & txtcampos(0).Text & "' " & Chr(13) & _
                      "Order By PEDPROVE.pp_pedido ASC, TFPRODUC.DESCRIPC ASC, TFPRODUC.CONTENID ASC"
     Dim RST As ADODB.Recordset
     Set RST = New ADODB.Recordset
     RST.Open "SELECT * FROM cattienda WHERE ticlave = '" & txtcampos(9).Text & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
     If Not (RST.EOF And RST.BOF) Then
        crpt.Formulas(4) = "ENTREGAR = 'PEDIDO POR PROVEEDOR A ENTREGAR EN BODEGA " & RST!tidescrip & "'"
        crpt.Formulas(5) = "DOMICILIO = '" & RST!Direccion & "     TEL: " & RST!telefonos & "'"
     End If
 Case 2 'Mixto
      If Mid(txtcampos(0).Text, 1, 3) = "C33" Or Mid(txtcampos(0).Text, 1, 3) = "C37" Then
         cARcRpt = App.Path & "\Pedprovm2.rpt"
      Else
         cARcRpt = App.Path & "\Pedprovm.rpt"
      End If
      crpt.SQLQuery = "SELECT PEDIDOS.p_pedido, PEDIDOS.p_proveedor, PEDIDOS.p_pedproveedor, " & _
                              "DETALLEFACTURA.df_costo, DETALLEFACTURA.df_cantsol, DETALLEFACTURA.df_cantsolp, DETALLEFACTURA.dg_decto1, DETALLEFACTURA.dg_decto2, DETALLEFACTURA.dg_decto3, DETALLEFACTURA.dg_decto4, DETALLEFACTURA.dg_decto5, DETALLEFACTURA.dg_decto6, DETALLEFACTURA.dg_iva, DETALLEFACTURA.dg_ieps, DETALLEFACTURA.dg_prelista, DETALLEFACTURA.dg_efectivo, " & _
                              "CATPROV.NOMPROVE, PEDPROVE.pp_observa, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC, TFPRODUC.clavedelprov " & _
                      "FROM pitico.dbo.PEDIDOS PEDIDOS, pitico.dbo.DETALLEFACTURA DETALLEFACTURA, pitico.dbo.CATPROV CATPROV, " & _
                            "pitico.dbo.PEDPROVE PEDPROVE, " & _
                            "pitico.dbo.TFPRODUC TFPRODUC " & _
                      "WHERE PEDIDOS.p_pedido = DETALLEFACTURA.df_pedido AND " & _
                            "PEDIDOS.p_proveedor = CATPROV.PROVE AND PEDIDOS.p_pedproveedor = PEDPROVE.pp_pedido AND " & _
                            "DETALLEFACTURA.df_prod = TFPRODUC.CONSEC AND " & _
                            "PEDIDOS.p_pedproveedor = '" & Trim(txtcampos(0).Text) & "' AND PEDIDOS.p_surtbodega = 1 AND PEDIDOS.p_cancelado = 0"
 Case 3 'Pedido de tienda
     'por la clave especial de colgate
     If Mid(txtcampos(0).Text, 1, 3) = "C33" Or Mid(txtcampos(0).Text, 1, 3) = "C37" Then
        cARcRpt = App.Path & "\PeCapCon2.rpt"
     Else
        cARcRpt = App.Path & "\PeCapCon1.rpt"
     End If
     ccondrpt = "FORMSELEC = {PEDIDOS.p_pedido} = '" & dbgrdPedsol.Columns(0).Text & "'"
 End Select
 crpt.WindowTitle = "Pedido por proveedor numero " & txtcampos(0).Text
 crpt.ReportFileName = cARcRpt
 crpt.Formulas(0) = ccondrpt
 If Index = 3 Then 'Pedido Por tienda
    crpt.WindowTitle = "Pedido sugerido numero " & dbgrdPedsol.Columns(0).Text
    crpt.Formulas(0) = ccondrpt
    crpt.Formulas(1) = "PEDIDO = 'PEDIDO SUGERIDO NUMERO [ " & dbgrdPedsol.Columns(0).Text & " ]'"
    crpt.Formulas(2) = ""
    crpt.Formulas(3) = ""
    crpt.Formulas(4) = ""
    crpt.SQLQuery = "SELECT PEDIDOS.p_proveedor, PEDIDOS.p_fecped, PEDIDOS.p_sucursal, PEDIDOS.p_fecentreal, PEDIDOS.p_observaciones, PEDIDOS.p_fecconfirma, " & _
                            "DETALLEFACTURA.df_costo, DETALLEFACTURA.df_cantsol, DETALLEFACTURA.Df_cantsolp, DETALLEFACTURA.dg_cajas, DETALLEFACTURA.dg_encajas, " & _
                            "CATPROV.NOMPROVE, CATTIENDA.tidescrip, TFPRODUC.DESCRIPC, TFPRODUC.CONTENID, TFPRODUC.MEDIDA, TFPRODUC.PAQUETES, TFPRODUC.CONSEC " & Chr(13) & _
                    "FROM pitico.dbo.PEDIDOS PEDIDOS, " & _
                          "pitico.dbo.DETALLEFACTURA DETALLEFACTURA, " & _
                          "pitico.dbo.CATPROV CATPROV, " & _
                          "pitico.dbo.CATTIENDA CATTIENDA, " & _
                          "pitico.dbo.TFPRODUC TFPRODUC " & Chr(13) & _
                    "WHERE PEDIDOS.p_pedido = DETALLEFACTURA.df_pedido AND " & _
                          "PEDIDOS.p_proveedor = CATPROV.PROVE AND " & _
                          "PEDIDOS.p_sucursal = CATTIENDA.ticlave AND " & _
                          "DETALLEFACTURA.df_prod = TFPRODUC.CONSEC AND " & _
                          "PEDIDOS.p_pedido = '" & AdoPedSol.Recordset!Folio & "' " & Chr(13) & _
                    "ORDER BY PEDIDOS.p_proveedor ASC, " & _
                          "TFPRODUC.DESCRIPC ASC, " & _
                          "TFPRODUC.CONTENID ASC"
 Else
    crpt.Formulas(1) = "FECELAB = 'FECHA DE ELABORACION:  " & txtcampos(3).Text & "'"
    crpt.Formulas(2) = "NUMPED = 'NUMERO DE PEDIDO:  " & txtcampos(0).Text & "'"
    crpt.Formulas(3) = "FECCONF = 'FECHA DE CONFIRM.:  " & IIf(chkCampos(0).Value, txtcampos(4).Text, "") & "' "
    'crpt.Formulas(4) = ""
 End If
 
 Set rsttemp = New ADODB.Recordset
 rsttemp.Open "SELECT * FROM Usuarios,Catprov WHERE Usuarios.Clave = Catprov.Comprador AND Prove = '" & txtcampos(1).Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
 If rsttemp.RecordCount > 0 Then crpt.Formulas(5) = "****** "
 crpt.Action = 1
 StbMensajes.SimpleText = cMensaje
 StbMensajes.Refresh
 Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdRptePed_Click()
   cMensaje = StbMensajes.SimpleText
   StbMensajes.SimpleText = Space(90) + "Espere un momento generando reporte..."
   StbMensajes.Refresh
   crpt.Connect = cCadConex
   crpt.ReportFileName = App.Path & "\PeCapCon.rpt"
   crpt.WindowTitle = "Pedido sugerido numero " & dbgrdPedsol.Columns(0).Text
   crpt.Formulas(0) = "FORMSELEC = '" & dbgrdPedsol.Columns(0).Text & "'"
   crpt.Formulas(1) = "PEDIDO = 'PEDIDO SUGERIDO NUMERO [ " & dbgrdPedsol.Columns(0).Text & " ]'"
   crpt.Formulas(2) = ""
   crpt.Formulas(3) = ""
   crpt.Formulas(4) = ""
   crpt.Formulas(5) = ""
   crpt.Action = 1
   StbMensajes.SimpleText = cMensaje
   StbMensajes.Refresh

End Sub


Private Sub cmdVerPedTie_Click()
If Not AdoPedProve.Recordset!pp_pedind Then
  dbgrdPedpro.Visible = False
  PicBotones.Visible = False
  fraPedTie.Visible = True
  'cmdReporte(3).SetFocus
Else
  nOp = 1 'para inicarle que el desp- fue llamado de Ped. indirecto
  TabInv = "INVENTARIO" & IIf(Trim(txtcampos(9).Text) = "16", "", Trim(txtcampos(9).Text))
  cn.Execute "DELETE FROM " & Tabla
  cCadena = "INSERT INTO " & Tabla & " (descripc, contenid, medida, consec,cantidad1,cantidad2, cantidad3, total, promedio, cantsol, paquetes, invactual, ininicial, instock, cantsolp, minimo, maximo,promocion,incant) " & _
            "SELECT p.descripc, p.contenid, LTRIM(STR(p.paquetes)) + ' X ' +  lTrim(str(p.contenid,10,3)) + space(2) + p.medida  AS MEDIDA, p.consec , cantidad1 = 0, cantidad2 = 0, cantidad3=0," & _
               "total = 0, promedio = 0, cantsol = 0, p.paquetes, i.incant, i.incant, i.instock, cantsolp = 0, i.minimo, i.maximo, ltrim(str(p.cajas)) + '/' + ltrim(str(encajas)), i.incant  " & _
               " FROM tfproduc p," & TabInv & " i WHERE activo = 1 and claprove = '" & txtcampos(1).Text & "' AND p.consec *= i.inprod"
  cn.Execute cCadena
  cn.Execute "UPDATE " & Tabla & " SET cantsol = dg_cantsol, cantsolp = dg_cantsolp FROM detalleglobal WHERE dg_producto = consec AND DG_PEDIDO = '" & Trim(txtcampos(0).Text) & "'"
  
  strcveprod = txtcampos(0).Text
    frmdespla.dbgrdTend.Caption = Me.txtcampos(9).Text & "  GENERAR PEDIDO INDIRECTO CON FOLIO " & txtcampos(0).Text
    frmdespla.dbgrdTend.Splits(1).Columns(9).Visible = True
    frmdespla.dbgrdTend.Splits(1).Columns(10).Visible = True
    frmdespla.dbgrdTend.Splits(1).Columns(3).Visible = False
    frmdespla.dbgrdTend.Splits(1).Columns(4).Visible = False
    frmdespla.dbgrdTend.Splits(1).Columns(5).Visible = False
    frmdespla.dbgrdTend.Splits(1).Columns(6).Visible = False
    frmdespla.dbgrdTend.Splits(1).Columns(12).Visible = False
    frmdespla.dbgrdTend.Splits(1).Columns(7).Visible = False
    frmdespla.dbgrdTend.Splits(1).Columns(13).Visible = False
    frmdespla.Show
End If
End Sub

Private Sub Command1_Click()
nOp = 0
cModo = "CAPTURARPEDIDO"
DEDONDE = "PROVEEDOR"
frmCaptPed.Show
SendKeys AdoPedSol.Recordset!Folio
SendKeys vbTab
End Sub

Private Sub dbgrdPedpro_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo Error:
If lNvoPedprove Then  'No se permite modificar cuando es un nuevo pedido
   MsgBox "HA OCURRIDO UN ERROR EN NUEVO PEDIDO Y SE ACTUALIZARON DATOS EN EL PEDIDO POR PROVEEDOR" & Chr(13) & "A CONTINUACION SE CANCELARAN LOS CAMBIOS", vbExclamation
   Cancel = True
Else
   marca = AdoPedpro.Recordset.Bookmark
   If UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "SOLXCAJA" Then
      cn.Execute "UPDATE detalleglobal SET dg_cantsol = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE dg_producto = '" & dbgrdPedpro.Columns(0).Text & "' AND dg_pedido = '" & txtcampos(0).Text & "'"
   ElseIf UCase(dbgrdPedpro.Columns(ColIndex).DataField) = "SOLXPZA" Then
      cn.Execute "UPDATE detalleglobal SET dg_cantsolP = " & dbgrdPedpro.Columns(ColIndex).Text & " WHERE dg_producto = '" & dbgrdPedpro.Columns(0).Text & "' AND dg_pedido = '" & txtcampos(0).Text & "'"
   End If
   Cancel = True
   Command2_Click
   dbgrdPedpro.Columns(0).Locked = True
   dbgrdPedpro.Columns(1).Locked = True
   dbgrdPedpro.Columns(2).Locked = True
   dbgrdPedpro.Columns(3).Locked = True
   AdoPedpro.Recordset.Bookmark = marca
   SendKeys "{DOWN}": SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}"
   SendKeys Chr(32): SendKeys "{ESC}"
End If
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Error:
 If KeyCode = 119 Then frmCalc.Show
 If KeyCode = 27 Then Unload Me
 If KeyCode = 115 Then
    PicBotones.Visible = False
    FraHist.Visible = True
    cMens = StbMensajes.Panels(1).Text
    StbMensajes.Panels(1).Text = "Espere un momento, consultando historial del producto"
    StbMensajes.Refresh
    Me.FraHist.Caption = AdoPedpro.Recordset!clave & " " & AdoPedpro.Recordset!descripcion & " " & AdoPedpro.Recordset!Present
    AdoEntPend.ConnectionString = cCadConex
    AdoEntPend.CommandType = adCmdText
    AdoEntPend.RecordSource = "SELECT * FROM Pedprove, detalleglobal WHERE dg_pedido = pp_pedido AND pp_recibe = 0 AND DG_PRODUCTO = '" & AdoPedpro.Recordset!clave & "' AND dg_cantsol > 0 AND pp_sucursal = '" & Trim(Mid(cSucursal, 1, 3)) & "' ORDER BY pp_fecConfirma DESC"
    AdoEntPend.Refresh
    AdoEntSur.CursorType = adOpenStatic
    AdoEntSur.ConnectionString = cCadConex
    AdoEntSur.CommandType = adCmdText
    AdoEntSur.RecordSource = "SELECT * FROM Pedprove, detalleglobal WHERE dg_pedido = pp_pedido AND pp_recibe = 1 AND DG_PRODUCTO = '" & AdoPedpro.Recordset!clave & "' AND dg_cantsol > 0 AND pp_sucursal = '" & Trim(Mid(cSucursal, 1, 3)) & "' ORDER BY pp_fecrecibe DESC"
    AdoEntSur.Refresh
    StbMensajes.Panels(1).Text = cMens
    StbMensajes.Refresh
 End If
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
Dim rsttemp  As ADODB.Recordset
'On Error GoTo Error:
Tabla = "TPEDSUG" & Trim(Mid(cCveDesUsu, 1, 4))  'Existe una tabla temporal por cada usuario

AdoPedProve.ConnectionString = cCadConex
AdoPedProve.CommandType = adCmdText
'Cuando se inicia es decir que no existe ningun pedido
If Not (frmpedBod.AdoPedidos.Recordset.BOF = True And frmpedBod.AdoPedidos.Recordset.EOF) Then
   AdoPedProve.RecordSource = "SELECT * FROM Pedprove WHERE pp_pedido = '" & frmpedBod.AdoPedidos.Recordset!pp_pedido & "'"
Else
   AdoPedProve.RecordSource = "SELECT * FROM Pedprove WHERE pp_pedido = 'XXX-X'"
End If
AdoPedProve.Refresh
'MsgBox AdoPedProve.Recordset!pp_pedind
lMove = True
If nOp = 1 Then  'Nuevo
     lblCajPza.Visible = False
     AdoPedProve.Recordset.AddNew
     AdoPedProve.Recordset!pp_pedind = False
     AdoProv.ConnectionString = cCadConex
     AdoProv.CommandType = adCmdText
     AdoProv.RecordSource = "SELECT * FROM cattienda"
     AdoProv.Refresh
     Do While Not AdoProv.Recordset.EOF
        cmbtienda.AddItem AdoProv.Recordset!tidescrip
        AdoProv.Recordset.MoveNext
     Loop
     AdoProv.Recordset.Close
     AdoProv.RecordSource = "SELECT * FROM Catprov WHERE activo = 1"
     AdoProv.Refresh
     chkCampos(0).Value = 0: chkCampos(1).Value = 0
     cmbProv.Clear
     AdoProv.Recordset.MoveFirst
     Do While Not AdoProv.Recordset.EOF
        If Not IsNull(AdoProv.Recordset!NOMPROVE) Then cmbProv.AddItem AdoProv.Recordset!NOMPROVE
        AdoProv.Recordset.MoveNext
     Loop
ElseIf cModo = "RECIBIR" Then
    lMove = True 'Bandera que no hace el scroll al grabar si no se cicla
    rsttemp.Open "SELECT SUM(DG_CANTSOL) AS TOTCAJ, SUM(DG_CANTSOLP) AS TOTPZA FROM DETALLEGLOBAL WHERE dg_pedido = '" & frmpedBod.AdoPedidos.Recordset!pp_pedido & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
    lblCajPza.Caption = "CAJAS:  " + CStr(rsttemp!totcaj) + Space(10) + "PIEZAS:  " + CStr(rsttemp!totpza)
End If
'Set rsttemp = New ADODB.Recordset
'se validan opciones por departamento
cmdGrabar.Enabled = (Nivel <> "P")
Me.cmdCancPedprove.Enabled = (Nivel <> "P")
Me.cmdActPreTie.Enabled = (Nivel <> "P")
Me.cmdCondCom.Enabled = (Nivel <> "P")
Me.cmdCancelar.Enabled = (Nivel <> "P")
Me.cmdAjuPed.Enabled = (Nivel <> "P")
Me.Command1.Enabled = (Nivel <> "P")
Me.dbgrdPedpro.AllowUpdate = (Nivel <> "P")
Me.dbgrdPedsol.AllowUpdate = (Nivel <> "P")
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmpedBod.Show
 If Not (frmpedBod.AdoPedidos.Recordset.BOF And frmpedBod.AdoPedidos.Recordset.EOF) Then
    frmpedBod.AdoPedidos.Recordset.MoveNext
    If frmpedBod.AdoPedidos.Recordset.EOF Then frmpedBod.AdoPedidos.Recordset.MoveLast
 End If
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
'Es el campo de observaciones y al teclear ";" inserta un caracter de retorno de linea
'Para que en le reporte salga en varias lineas
If Index = 5 And KeyAscii = 59 Then
   txtcampos(Index).Text = txtcampos(Index).Text + Chr(13)
   KeyAscii = 0
   SendKeys "^{END}"
End If

End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
Dim rstTie As ADODB.Recordset
Dim rsttemp As ADODB.Recordset
Dim cCadena As String
Dim nfol As Integer
On Error GoTo Error:
Select Case Index
Case 1 'Clave del proveedor
     txtcampos(Index).Text = Trim(UCase(txtcampos(Index).Text))
     txtcampos(Index).Refresh
     'NUEVO = Ver si existen pedidos para generar un pedido por proveedor global
     If nOp = 1 Then
        If txtcampos(Index).Text = "" Or IsNull(txtcampos(Index).Text) Then
           cmbProv.SetFocus
           Exit Sub
        Else
           AdoProv.Recordset.MoveFirst
           AdoProv.Recordset.Find "Prove= '" & Trim(txtcampos(Index).Text) & "'"
           If AdoProv.Recordset.EOF = True Then
               cmbProv.SetFocus
               Exit Sub
           End If
        End If
        txtcampos(6).Text = date + AdoProv.Recordset!frecuencia
        chkCampos(0).Visible = True
        cmbProv.Text = AdoProv.Recordset!NOMPROVE
        'Obtengo el folio del pedido en base al proveedor Especif.
        Set rsttemp = New ADODB.Recordset
        rsttemp.ActiveConnection = cCadConex
        rsttemp.CursorType = adOpenKeyset
        rsttemp.LockType = adLockOptimistic
        rsttemp.Source = "SELECT pp_pedido FROM Pedprove WHERE PP_proveedor = '" & txtcampos(1).Text & "'"
        rsttemp.Open
        nFolMay = 0
        While Not rsttemp.EOF
             nfol = Mid(rsttemp!pp_pedido, InStr(1, rsttemp!pp_pedido, "-") + 1)
             'No se consideran las llegadas de carbonera que son mayores a 1000
             nFolMay = IIf(nFolMay < nfol And nfol < 1000, nfol, nFolMay)
             rsttemp.MoveNext
        Wend
        
        If nFolMay = 0 Then
           ' txtcampos(0).Text = Mid(Trim(Mid(txtcampos(1).Text, 1, 3)), 1, 3) + "-1"
            txtcampos(0).Text = Trim(txtcampos(1).Text) + "-1"
        Else
            txtcampos(0).Text = txtcampos(1).Text + "-" + Trim(Str(nFolMay + 1))
        End If
        
        If lpprov Then   'Es un pedido indirecto
           cn.Execute "DELETE FROM " & Tabla
           TabInv = "INVENTARIO" & IIf(Trim(txtcampos(9).Text) = Trim(Mid(cSucursal, 1, 3)), "", Trim(txtcampos(9).Text))
           cCadena = "INSERT INTO " & Tabla & " (descripc, contenid, medida, consec,cantidad1,cantidad2, cantidad3, total, promedio, cantsol, paquetes, invactual, ininicial, instock, cantsolp, minimo, maximo,promocion,incant) " & _
              "SELECT p.descripc, p.contenid, LTRIM(STR(p.paquetes)) + ' X ' +  lTrim(str(p.contenid,10,3)) + space(2) + p.medida  AS MEDIDA, p.consec , cantidad1 = 0, cantidad2 = 0, cantidad3=0," & _
              "total = 0, promedio = 0, cantsol = 0, p.paquetes, i.incant, i.incant, i.instock, cantsolp = 0, i.minimo, i.maximo, ltrim(str(p.cajas)) + '/' + ltrim(str(encajas)), i.incant  " & _
              " FROM tfproduc p," & TabInv & " i WHERE activo = 1 and claprove = '" & txtcampos(1).Text & "' AND p.consec *= i.inprod"
           cn.Execute cCadena
           strcveprod = txtcampos(0).Text
           frmdespla.dbgrdTend.Caption = Trim(Me.txtcampos(9).Text) & "  - " & "GENERAR PEDIDO INDIRECTO CON FOLIO " & txtcampos(0).Text
           Unload frmPedProv
           frmdespla.dbgrdTend.Splits(1).Columns(9).Visible = True
           frmdespla.dbgrdTend.Splits(1).Columns(10).Visible = True
           frmdespla.dbgrdTend.Splits(1).Columns(3).Visible = False
           frmdespla.dbgrdTend.Splits(1).Columns(4).Visible = False
           frmdespla.dbgrdTend.Splits(1).Columns(5).Visible = False
           frmdespla.dbgrdTend.Splits(1).Columns(6).Visible = False
           frmdespla.dbgrdTend.Splits(1).Columns(12).Visible = False
           frmdespla.dbgrdTend.Splits(1).Columns(7).Visible = False
           frmdespla.dbgrdTend.Splits(1).Columns(13).Visible = False
           frmdespla.Show 1
           Exit Sub
        End If
        
        cCond = "p_situacion = 0 AND p_cancelado = 0 AND p_proveedor = '" & txtcampos(1).Text & "'"
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
        lbletiquetas(4).Visible = True
        txtcampos(5).Visible = True
        txtcampos(5).SetFocus
        cmbProv.Enabled = False
        dbgrdPedpro.Visible = True

        lbletiquetas(5).Visible = True
        txtcampos(6).Visible = True
        cmdCancPedprove.Enabled = False
        cmdReporte(0).Enabled = False
        cmdReporte(1).Enabled = False
        cmdActPreTie.Enabled = False
        PicBotones.Visible = True
        
     Else 'Si es consulta del pedido confirmado O recepcion
        cCond = "p_pedproveedor = '" & frmpedBod.AdoPedidos.Recordset!pp_pedido & "' AND P_cancelado = 0"
        For N = 0 To 4
           If N < 4 Then lbletiquetas(N).Visible = True
           txtcampos(N).Visible = True
           txtcampos(N).Enabled = False
        Next
        cmbProv.Visible = True
        cmbProv.Text = frmpedBod.cmbProved.Text
        chkCampos(0).Visible = True
        chkCampos(1).Visible = True
        chkCampos(0).Enabled = (txtcampos(1).Text = "JAR")
        chkCampos(1).Enabled = cModo = "RECIBIR"
        dbgrdPedpro.Visible = True
        dbgrdPedpro.Width = ScaleWidth - 400
        lbletiquetas(6).Visible = True
        txtcampos(7).Visible = True
        lbletiquetas(7).Visible = True
        txtcampos(8).Visible = True
        If cModo = "RECIBIR" Then
           cmdReporte(0).Visible = AdoPedProve.Recordset!pp_recibe
           cmdReporte(1).Visible = AdoPedProve.Recordset!pp_recibe
           cmdReporte(3).Visible = AdoPedProve.Recordset!pp_recibe
           cmdReporte(2).Visible = AdoPedProve.Recordset!pp_recibe
        Else
           cmdReporte(0).Visible = True
           cmdReporte(1).Visible = True
           cmdReporte(3).Visible = True
           cmdReporte(2).Visible = True
           cmdVerPedTie.Visible = True
        End If
        AdoFacturas.ConnectionString = cCadConex
        AdoFacturas.CommandType = adCmdText
        AdoFacturas.RecordSource = "SELECT * FROM [Facturas] WHERE [f_pedido] = '" & txtcampos(0).Text & "'"
        AdoFacturas.Refresh
        If cModo = "RECIBIR" Then
            dbgrdPedpro.Splits(0).Locked = True
            If Not AdoPedProve.Recordset!pp_recibe Then
               cn.Execute "UPDATE DetalleGlobal SET dg_cantRec = dg_cantidad"
               If AdoFacturas.Recordset.EOF = True Then AdoFacturas.Recordset.AddNew 'Para que no se borren los datos al ecribir en los campos del control adofacturas
               cmdGrabar.Visible = True
               cmdAjuPed.Visible = True
               cmdCondCom.Visible = True
            Else
               For N = 0 To 3
                  txtRecib(N).Visible = True
                  txtRecib(N).Enabled = False
               Next
               chkCampos(1).Enabled = False
               txtRecib(0).Text = AdoPedProve.Recordset!pp_fecrecibe
               'dbgrdRec.AllowUpdate = False
               cmdGrabar.Visible = True
               cmdAjuPed.Visible = True
               cmdGrabar.Enabled = True
               cmdAjuPed.Enabled = True
               cmdCondCom.Visible = True
               'cmdCondCom.Enabled = False
            End If
            AdoDetGlo.ConnectionString = cCadConex
            AdoDetGlo.CommandType = adCmdText
            AdoDetGlo.LockType = adLockOptimistic
            AdoDetGlo.RecordSource = "SELECT * FROM DetalleGlobal WHERE dg_pedido = '" & txtcampos(0).Text & "' ORDER BY dg_producto"
            AdoDetGlo.Refresh
            'dbgrdRec.Visible = True
            chkCampos(1).Visible = True
        'Si es opcion ver confirmado y ya es recibido
        ElseIf AdoPedProve.Recordset!pp_recibe Then
            For N = 0 To 3
                txtRecib(N).Visible = True
                txtRecib(N).Enabled = False
            Next
            cmdGrabar.Visible = True
            cmdGrabar.Enabled = False
            cmdAjuPed.Visible = True
            cmdAjuPed.Enabled = False
            cmdGrabar.Enabled = True
            cmdGrabar.Enabled = False
            cmdCondCom.Visible = True
            'CmdExporta.Enabled = False
            txtcampos(5).Enabled = False
            dbgrdPedpro.AllowUpdate = False
        'Si es opcion ver confirmado y no ha sido recibido
        Else
            cmdGrabar.Visible = True
            cmdAjuPed.Visible = True
            cmdGrabar.Enabled = True
            cmdAjuPed.Enabled = True
            cmdCondCom.Visible = True
            txtcampos(5).Enabled = True
        End If
        lbletiquetas(4).Visible = True
        txtcampos(5).Visible = True
        lbletiquetas(5).Visible = True
        txtcampos(6).Visible = True
        
        Set rsttemp = New ADODB.Recordset
        
        rsttemp.Open "SELECT COUNT(DG_PRODUCTO) AS TOTPRO, SUM(DG_CANTSOL) AS TOTCAJ, SUM(DG_CANTSOLP) AS TOTPZA FROM DETALLEGLOBAL WHERE dg_pedido = '" & txtcampos(0).Text & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
        If rsttemp!totpro > 0 Then
           lblCajPza.Visible = True
           lblCajPza.Caption = "PRODUCTOS: " & CStr(rsttemp!totpro) & Space(5) & "CAJAS: " + CStr(rsttemp!totcaj) + Space(5) + "PIEZAS: " + CStr(rsttemp!totpza)
        End If
        PicBotones.Visible = True

        'si es enviado desactivar botones
        'MsgBox AdoPedProve.Recordset!pp_enviado
        If Trim(AdoPedProve.Recordset!pp_enviado) = "1" Then
            cmdCondCom.Enabled = False
            cmdAjuPed.Enabled = False
            cmdGrabar.Enabled = False
            Command1.Enabled = False
            cmdCancelar.Enabled = False
            dbgrdPedpro.AllowUpdate = False
            dbgrdPedsol.AllowUpdate = False
         End If
     End If
        rsttemp.Close
     If cModo <> "RECIBIR" Then
        If AdoPedProve.Recordset!pp_pedind Then
           AdoPedpro.ConnectionString = cCadConex
           'AdoPedpro.LockType = adLockOptimistic
           AdoPedpro.CommandType = adCmdText
            AdoPedpro.RecordSource = "SELECT dg_producto AS CLAVE, descripc AS DESCRIPCION, str(paquetes) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT, dg_cantsol AS SOLxCAJ, dg_cantsolp AS SOLxPZA FROM detalleglobal,tfproduc WHERE dg_producto = consec AND dg_pedido = '" & AdoPedProve.Recordset!pp_pedido & "' ORDER BY descripc"
            AdoPedpro.Refresh
            cmdReporte(1).Enabled = True
            cmdCondCom.Enabled = True
            cmdActPreTie.Enabled = False
            cmdGrabar.Enabled = True
            lpprov = True  'Para indicar que es un pedido indirecto
            Set rstDesPro = New ADODB.Recordset
            rstDesPro.Open "SELECT * FROM Detalleglobal,TFPRODUC, DESCPROD WHERE dg_pedido = '" & txtcampos(0).Text & "' AND DETALLEGLOBAL.dg_producto = TFPRODUC.CONSEC AND TFPRODUC.CONSEC = DESCPROD.PRODUCTO AND DETALLEGLOBAL.DG_PRODUCTO = DESCPROD.PRODUCTO ", cn, adOpenKeyset, adLockOptimistic, adCmdText
            'MsgBox rstDesPro.RecordCount
      Else
        cmen = StbMensajes.SimpleText
        StbMensajes.SimpleText = Space(55) & "Espere un momento verificando pedidos pendientes de confirmar."
        StbMensajes.Refresh
      If cModo = "" Then
        'Para crear ref. cruzada Sql Server recorro todas las tiendas
        Set rstTie = New ADODB.Recordset
        rstTie.Source = "SELECT * FROM cattienda WHERE not Prioridad is Null ORDER BY Prioridad"
'        MsgBox cCadConex
        rstTie.ActiveConnection = cCadConex
        rstTie.LockType = adLockOptimistic
        rstTie.Open
        'Genera la cadena del origen de datos es una referencia cruzada y
        'utiliza una vista de Sql. (DetPedTie)
        cCadena = "SELECT df_prod AS CLAVE, descripc As DESCRIPCION, str(paquetes) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT, max(incant) as EXISTENCIA ," + Chr(13) _
        & " SUM(DF_CANTSOL) As SOLxCAJA, SUM(df_cantsolp) AS SOLxPZA,"
        'MsgBox cCadena
        While Not rstTie.EOF
            'Claves de las sucursales que enviaran sus pedidos
            'If InStr(1, "10-1-2-78-8-4-21-20-22-12-5-16-15-14-13-17-18-19", Trim(rstTie!ticlave) + "-") > 0 Then
            cCadena = cCadena + Chr(13) + " SUM(CASE p_sucursal WHEN '" & Trim(rstTie!ticlave) & "' THEN df_cantsol ELSE 0 END) AS " & Trim(Mid(rstTie!tidescrip, 1, 5)) & "xCAJ ," + " SUM(CASE p_sucursal WHEN '" & Trim(rstTie!ticlave) & "' THEN df_cantsolp ELSE 0 END) AS " & Trim(Mid(rstTie!tidescrip, 1, 5)) & "xPZA ,"
            'End If
            'MsgBox cCadena
            rstTie.MoveNext
        Wend
     '   MsgBox cCadena
        cCadena = Mid(cCadena, 1, Len(cCadena) - 2) & Chr(13) _
        & " From PedTie1 WHERE " & cCond _
        & " GROUP BY df_prod, descripc, str(paquetes) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA ORDER BY Descripcion"
      Else
         cCadena = "SELECT dg_producto AS CLAVE, descripc As DESCRIPCION, str(paquetes) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT, InCant AS EXISTENCIA,dg_cantsol AS SOLxCAJA, dg_cantsolp AS SOLxPZA FROM detalleglobal,tfproduc,Inventario WHERE dg_producto = consec AND dg_producto = Inprod AND dg_pedido = '" & txtcampos(0).Text & "' ORDER BY descripc,contenid"
      End If
        'Obtengo cantidades solicitadas por tienda de un proveedor especificado
        AdoPedpro.CursorType = adOpenKeyset
        AdoPedpro.LockType = adLockOptimistic
        AdoPedpro.ConnectionString = cCadConex
        AdoPedpro.CommandType = adCmdText
        AdoPedpro.RecordSource = cCadena
        AdoPedpro.Refresh
        lMove = False
        If cModo <> "" And AdoPedProve.Recordset!pp_enviado <> "1" Then
           dbgrdPedpro.Columns(0).Locked = True
           dbgrdPedpro.Columns(1).Locked = True
           dbgrdPedpro.Columns(2).Locked = True
           dbgrdPedpro.AllowUpdate = True
        End If
        
        'Obtengo los folios de pedidos que solicitaron las tiendas
        AdoPedSol.LockType = adLockOptimistic
        AdoPedSol.CursorType = adOpenDynamic
        AdoPedSol.ConnectionString = cCadConex
        AdoPedSol.CommandType = adCmdText
        If cModo = "VERCONF" Then
            AdoPedSol.RecordSource = "SELECT p_pedido as FOLIO, tidescrip AS SUCURSAL, P_fecPed As FECHA_SOL, P_surtbodega as SURTBOD FROM PEDIDOS,CATTIENDA WHERE P_SUCURSAL = TICLAVE AND P_PEDPROVEEDOR = '" & txtcampos(0).Text & "' AND p_cancelado = 0 ORDER BY SUCURSAL"
            AdoPedSol.Refresh
        Else 'Si esta en nuevo pedido
            AdoPedSol.RecordSource = "SELECT DISTINCT p_pedido as FOLIO, tidescrip AS SUCURSAL, P_fecPed As FECHA_SOL, P_surtbodega as SURTBOD FROM PedDetTie WHERE " & cCond & " ORDER BY SUCURSAL"
            AdoPedSol.Refresh
        End If
        If AdoPedSol.Recordset.BOF And AdoPedSol.Recordset.EOF Then
           MsgBox "NO EXISTEN PEDIDOS DE TIENDA PARA CONFIRMAR" & Chr(13) & "DEL PROVEEDOR ESPECIFICADO", vbExclamation
        End If
        cmdReporte(3).Visible = True
        cmdReporte(2).Visible = True
        StbMensajes.SimpleText = Space(45) & "Calculando el total de piezas, cajas y buscando las promociones de los articulos... "
        StbMensajes.Refresh
        
        If Not (AdoPedpro.Recordset.BOF And AdoPedpro.Recordset.EOF) Then
            AdoPedpro.Recordset.MoveFirst
            nCaj = 0: nPza = 0: npro = 0
            While Not AdoPedpro.Recordset.EOF
                nCaj = nCaj + AdoPedpro.Recordset!SolxCaja
                nPza = nPza + AdoPedpro.Recordset!SolxPza
                npro = npro + 1
                AdoPedpro.Recordset.MoveNext
            Wend
            lMove = True
            lblCajPza.Visible = True
            If IsNull(nPza) Then nPza = 0
            If IsNull(nCaj) Then nCaj = 0
            lblCajPza.Caption = "PRODUCTOS: " & CStr(npro) & Space(5) & "CAJAS : " & CStr(nCaj) & Space(5) & "PIEZAS: " & CStr(nPza)
        End If
        Set rstDesPro = New ADODB.Recordset
        rstDesPro.Open "SELECT * FROM PEDIDOS, DETALLEFACTURA,TFPRODUC, DESCPROD WHERE " & cCond & " AND PEDIDOS.P_PEDIDO = DETALLEFACTURA.DF_PEDIDO AND DETALLEFACTURA.DF_PROD = TFPRODUC.CONSEC AND TFPRODUC.CONSEC = DESCPROD.PRODUCTO AND DETALLEFACTURA.DF_PROD = DESCPROD.PRODUCTO ", cn, adOpenKeyset, adLockOptimistic, adCmdText
        StbMensajes.SimpleText = cmen
        StbMensajes.Refresh
      End If
     Else 'Si es recibir
        CADENA = "SELECT dg_producto AS CLAVE, LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, descripc AS DESCRIPCION, dg_cantidad AS CANT_SOL FROM DetalleGlobal,tfproduc WHERE dg_pedido = '" & txtcampos(0).Text & "' AND TFPRODUC.consec = DETALLEGLOBAL.dg_producto ORDER BY dg_producto"
        AdoPedpro.RecordSource = CADENA
        AdoPedpro.Refresh
        Me.dbgrdPedpro.Refresh
     End If
     If cModo = "RECIBIR" Then dbgrdPedpro.Columns(3).Width = 2150
     cmdregresar.Visible = True
     dbgrdPedpro.Columns(0).Locked = True
     dbgrdPedpro.Columns(1).Locked = True
     dbgrdPedpro.Columns(2).Locked = True
     dbgrdPedpro.Columns(3).Locked = True
     dbgrdPedpro.Columns(1).Width = 5560
     SendKeys vbTab
Case 9
     'If nOp = 1 Then
        If txtcampos(Index).Text = "" Or IsNull(txtcampos(Index).Text) Then
           cmbtienda.SetFocus
           Exit Sub
        Else
           Dim RST As ADODB.Recordset
           Set RST = New ADODB.Recordset
           RST.Open "SELECT * FROM cattienda WHERE ticlave = '" & txtcampos(Index).Text & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
           If RST.EOF And RST.BOF Then
               cmbtienda.SetFocus
               Exit Sub
           Else
              cmbtienda.Text = Trim(RST!tidescrip)
           End If
           RST.Close
           Set RST = Nothing
        End If
     'End If
     lbletiquetas(1).Visible = True
     txtcampos(1).Visible = True
     cmbProv.Visible = True
     'Me.dbgrdPedpro.Visible = False
     If txtcampos(1).Enabled Then txtcampos(1).SetFocus
End Select
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub txtRecib_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub

'Funcion que calcula el costo sin promocion de un producto.
Private Function CostSinProm(producto As String) As Double
Dim npreciodes As Double
Dim nprecio As Double
Dim i As Integer
Dim rs As ADODB.Recordset
On Error GoTo Error:
    
    Set rs = New ADODB.Recordset
    'Verifico si existe en descprod
    rs.Open "SELECT * FROM DESCPROD WHERE PRODUCTO = '" & Trim(producto) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rs.RecordCount > 0 Then
       nprecio = rs!preciolista
    Else
       MsgBox "LA CLAVE DEL PRODUCTO " & producto & " NO EXISTE EN LA TABLA DE DESCUENTOS" & Chr(13) & "Y NO SE LE ASIGNARA COSTO", vbCritical
       CostSinProm = 0
       Exit Function
    End If
    
    'calcula cargos %
        'For I = 2 To 3
        '    nprecio = nprecio + (nprecio * (Val(Txtcar(I).Text) / 100))
        'Next I
        'cargo efectivo
        'nprecio = nprecio + Val(Txtcar(4).Text)
    nprecio = nprecio + (nprecio * (Val(rs!cargo4)) / 100)
    nprecio = nprecio + (nprecio * (Val(rs!cargo3)) / 100)
    nprecio = nprecio + Val(rs!cargo5)
     
    'calcula descuentos %
    'nprecio = Round(nprecio, 2)
    nprecio = nprecio
        'For I = 0 To 3
        '     nprecio = nprecio - (nprecio * (Val(Txtdes(I).Text) / 100))
        'Next I
        'descuento efectivo
        'nprecio = nprecio - Val(Txtdes(4).Text)
    nprecio = nprecio - (nprecio * (Val(rs!decto1) / 100))
    nprecio = nprecio - (nprecio * (Val(rs!decto2) / 100))
    nprecio = nprecio - (nprecio * (Val(rs!decto3) / 100))
    nprecio = nprecio - (nprecio * (Val(rs!decto4) / 100))
    'Descuento efectivo
    nprecio = nprecio - Val(rs!financiero)
    
        'descuento financiero
        'nprecio = nprecio - (nprecio * (Val(Txtdes(5).Text) / 100))
        'Mskcompra.Text = Round(nprecio, 2)
    'Descuento financiero
    nprecio = nprecio - (nprecio * (Val(rs!efectivo) / 100))
    'Mskcompra.Text = Round(nprecio, 2)
    
    CostSinProm = Round(nprecio, 3)
Exit Function
Error:
    MsgBox Err.Description, vbCritical
End Function

