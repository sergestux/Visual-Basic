VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmHisto 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial del producto "
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   Icon            =   "frmHisto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtpza 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   22
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtpzateorico 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8280
      TabIndex        =   19
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtcajteorico 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtfecha 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtsalida 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   12
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtEntrada 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DatgEnt 
      Bindings        =   "frmHisto.frx":0442
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   12515
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
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
      Caption         =   "ENTRADAS REALIZADAS   "
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "cantrec"
         Caption         =   "CAJ.REC."
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
      BeginProperty Column01 
         DataField       =   "existencia"
         Caption         =   "EXIST. ANT."
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
      BeginProperty Column02 
         DataField       =   "folio"
         Caption         =   "     FOLIO"
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
         DataField       =   "fecharec"
         Caption         =   "FECHA REC."
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
         DataField       =   "fechaelab"
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
      BeginProperty Column05 
         DataField       =   "fechaconf"
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
      BeginProperty Column06 
         DataField       =   "cantsol"
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
      BeginProperty Column07 
         DataField       =   "facturas"
         Caption         =   "  FACTURA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "importe"
         Caption         =   "IMPORTE"
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
      BeginProperty Column09 
         DataField       =   "tipo"
         Caption         =   "TIPO ENT."
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
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1365.165
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DatGSal 
      Bindings        =   "frmHisto.frx":045C
      Height          =   7095
      Left            =   6480
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   12515
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
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
      Caption         =   "SALIDAS     [ TRASLADOS ]"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cantREC"
         Caption         =   "CAJ. ENV."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cantrecp"
         Caption         =   "PZA.ENV."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "fechaREC"
         Caption         =   "FECHA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "DD/MM/YY hh:MM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "sucursal"
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
      BeginProperty Column04 
         DataField       =   "folio"
         Caption         =   "FOLIO"
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
      BeginProperty Column05 
         DataField       =   "tipo"
         Caption         =   "TIPO"
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
         Locked          =   -1  'True
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1260.284
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtExi 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtIni 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin MSAdodcLib.Adodc AdoEnvios 
      Height          =   330
      Left            =   0
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
      Caption         =   "AdoEnvios"
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
   Begin MSAdodcLib.Adodc AdoEntradas 
      Height          =   330
      Left            =   2160
      Top             =   -120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoEntradas"
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
   Begin MSAdodcLib.Adodc AdoBack 
      Height          =   330
      Left            =   3960
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
      Caption         =   "AdoBack"
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
   Begin MSAdodcLib.Adodc AdoEntPend 
      Height          =   330
      Left            =   3960
      Top             =   480
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
   Begin MSAdodcLib.Adodc AdoPedidos 
      Height          =   330
      Left            =   3960
      Top             =   480
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
   Begin MSAdodcLib.Adodc AdoEntSur 
      Height          =   330
      Left            =   3840
      Top             =   480
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
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   9360
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "                                             Para salir sin guardar los cambios presione el boton regresar"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraPend 
      Caption         =   "Pendientes"
      Height          =   7455
      Left            =   0
      TabIndex        =   6
      Top             =   1280
      Visible         =   0   'False
      Width           =   11535
      Begin MSDataGridLib.DataGrid DatgEntPend 
         Bindings        =   "frmHisto.frx":0474
         Height          =   3615
         Left            =   360
         TabIndex        =   7
         Top             =   360
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
         ColumnCount     =   6
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "pp_observa"
            Caption         =   "    OBSERVACIONES"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1920.189
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   3734.929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DatgEntSurt 
         Bindings        =   "frmHisto.frx":048D
         Height          =   3375
         Left            =   360
         TabIndex        =   8
         Top             =   4080
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5953
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
         Caption         =   "ENTRADAS RECIBIDAS Y NO SURTIDAS AL 100 %    [ PEDIDOS ]"
         ColumnCount     =   7
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
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
         BeginProperty Column06 
            DataField       =   "pp_observa"
            Caption         =   "    OBSERVACIONES"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1920.189
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   3734.929
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  Cajas            Piezas"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   9480
      TabIndex        =   23
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "="
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cajas           Piezas"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7440
      TabIndex        =   20
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exis. Segun E/S"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha de Inv. Inicial"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbletiquetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salidas"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lbletiquetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entradas"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lbletiquetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Existencia en cajas"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   9960
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lbletiquetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Inventario Inicial"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmHisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DatgBac_KeyPress(KeyAscii As Integer)
If KeyAscii = 37 Then
   Unload Me
End If
End Sub

Private Sub DatgEnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
   nOp = 1
End If
End Sub

Private Sub DatgEnt_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 If DatgEnt.SelBookmarks.Count > 0 Then DatgEnt.SelBookmarks.Remove 0
 DatgEnt.SelBookmarks.Add DatgEnt.RowBookmark(DatgEnt.Row)
End Sub

Private Sub DatGSal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub DatGSal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 If DatGSal.SelBookmarks.Count > 0 Then DatGSal.SelBookmarks.Remove 0
 DatGSal.SelBookmarks.Add DatGSal.RowBookmark(Me.DatGSal.Row)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 119 Then frmCalc.Show 1 'F8
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
On Error GoTo Error:
Set rs = New ADODB.Recordset
rs.Open "SELECT SUM(Cantrec) AS CAJREC FROM histent WHERE entrada = 1", cn, adOpenKeyset, adLockOptimistic, adCmdText
frmHisto.txtEntrada.Text = IIf(IsNull(rs!Cajrec), 0, rs!Cajrec)
rs.Close
rs.Open "SELECT SUM(Cantrec) AS CAJSAL FROM histent WHERE entrada = 0", cn, adOpenKeyset, adLockOptimistic, adCmdText
'rs.Open "SELECT SUM(dt_cantidad) AS CAJSAL FROM Traslados, DetalleTraslado, Cattienda WHERE t_clave = dt_clave AND T_sucursalReceptor = ticlave AND t_motivocancela is null AND Dt_cantidad > 0 AND Dt_producto = '" & frmModInv.AdoModInv.Recordset!Inprod & "' AND t_entrada = 0 AND T_ENVIADO = 1 ", cn, adOpenKeyset, adLockOptimistic, adCmdText
frmHisto.txtsalida.Text = IIf(IsNull(rs!Cajsal), 0, rs!Cajsal)

frmHisto.AdoEnvios.ConnectionString = cCadConex
frmHisto.AdoEnvios.CommandType = adCmdText
frmHisto.AdoEnvios.RecordSource = "SELECT * FROM histent WHERE entrada = 0 ORDER BY fecharec DESC"
frmHisto.AdoEnvios.Refresh

frmHisto.AdoEntradas.ConnectionString = cCadConex
frmHisto.AdoEntradas.CommandType = adCmdText
frmHisto.AdoEntradas.RecordSource = "SELECT * FROM histent WHERE entrada = 1 ORDER BY fecharec DESC"
frmHisto.AdoEntradas.Refresh
Set rs = Nothing
frmHisto.Caption = "HISTORIAL DEL PRODUCTO: " & frmModInv.AdoModInv.Recordset!Inprod & frmModInv.AdoModInv.Recordset!descripc & " " & frmModInv.AdoModInv.Recordset!medida
txtIni.Text = frmModInv.AdoModInv.Recordset!InInicial
txtExi.Text = frmModInv.AdoModInv.Recordset!InCant
txtpza.Text = frmModInv.AdoModInv.Recordset!InCantPza
txtcajteorico.Text = Val(txtIni.Text) + Val(txtEntrada.Text) - Val(txtsalida.Text)
Exit Sub
Error:
MsgBox Err.Description
End Sub

