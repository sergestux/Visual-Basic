VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModInv 
   Caption         =   "Modificar Inventario"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   180
   ClientWidth     =   11880
   Icon            =   "frmModInv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraProv 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccione proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1800
      TabIndex        =   37
      Top             =   2640
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Regresar"
         Height          =   495
         Left            =   4200
         Picture         =   "frmModInv.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   495
         Left            =   2160
         Picture         =   "frmModInv.frx":0FB4
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Consultar existencias"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmbProved 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtClave 
         Height          =   285
         Left            =   840
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   735
      End
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
      ForeColor       =   &H00404040&
      Height          =   1695
      Left            =   3600
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton btnresp 
         Caption         =   "&Respaldo"
         Height          =   375
         Left            =   2880
         TabIndex        =   35
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ini"
         Height          =   375
         Left            =   1560
         TabIndex        =   34
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label lblContra 
         BackColor       =   &H00C0C000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame frapend 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   1320
      TabIndex        =   47
      Top             =   1200
      Visible         =   0   'False
      Width           =   9495
      Begin MSDataGridLib.DataGrid DBGRDPEND 
         Bindings        =   "frmModInv.frx":10B6
         Height          =   4575
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   16761024
         BorderStyle     =   0
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
         Caption         =   "NOMBRE DEL PRODUCTO Y PRESENTACION"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "cnombre"
            Caption         =   "Nombre"
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
            DataField       =   "folpreventa"
            Caption         =   "Pvta"
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
         BeginProperty Column03 
            DataField       =   "fecha"
            Caption         =   "fecha"
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
         BeginProperty Column04 
            DataField       =   "cantidad"
            Caption         =   "Cajas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "cantidadp"
            Caption         =   "Piezas"
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
            DataField       =   "importe"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   3569.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   959.811
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdpenreg 
         Caption         =   "Regresar"
         Height          =   255
         Left            =   8520
         TabIndex        =   49
         Top             =   80
         Width           =   855
      End
      Begin MSAdodcLib.Adodc AdoPend 
         Height          =   330
         Left            =   1680
         Top             =   4560
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
         Caption         =   "AdoPend"
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
   Begin VB.PictureBox PicBotones 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   11820
      TabIndex        =   14
      Top             =   7575
      Width           =   11880
      Begin VB.OptionButton optVerInv 
         BackColor       =   &H00808080&
         Caption         =   "Con existencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   8520
         TabIndex        =   44
         Top             =   420
         Width           =   1600
      End
      Begin VB.OptionButton optVerInv 
         BackColor       =   &H00808080&
         Caption         =   "Solo activos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   8520
         TabIndex        =   43
         Top             =   210
         Value           =   -1  'True
         Width           =   1600
      End
      Begin VB.OptionButton optVerInv 
         BackColor       =   &H00808080&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   8520
         TabIndex        =   42
         Top             =   0
         Width           =   1600
      End
      Begin VB.CommandButton cmdRpteInv 
         Caption         =   "&Inv."
         Height          =   480
         Left            =   6120
         Picture         =   "frmModInv.frx":10CC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Visualiza inventario del proveedor seleccionado"
         Top             =   80
         Width           =   700
      End
      Begin VB.CommandButton cmdprod 
         Caption         =   "&Prod."
         Height          =   480
         Left            =   5430
         Picture         =   "frmModInv.frx":15FE
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Visualizar características del producto"
         Top             =   70
         Width           =   700
      End
      Begin VB.CommandButton cmdBuscaDesc 
         Height          =   300
         Left            =   600
         Picture         =   "frmModInv.frx":1700
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Busqueda por descripcion"
         Top             =   300
         Width           =   500
      End
      Begin VB.CommandButton cmdBuscaCve 
         Height          =   300
         Left            =   120
         Picture         =   "frmModInv.frx":17FA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Busqueda por clave"
         Top             =   300
         Width           =   500
      End
      Begin VB.CommandButton cmdBuscaBarra 
         Height          =   300
         Left            =   600
         Picture         =   "frmModInv.frx":18F4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Busqueda por codigo de barras"
         Top             =   0
         Width           =   500
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   480
         Left            =   7620
         Picture         =   "frmModInv.frx":1A2A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Regresar al menu principal"
         Top             =   70
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Historial"
         Height          =   480
         Left            =   6900
         Picture         =   "frmModInv.frx":1B9C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Regresar al menu principal"
         Top             =   70
         Width           =   735
      End
      Begin VB.CommandButton cmdInvIni 
         Caption         =   "&Precios"
         Height          =   480
         Left            =   4740
         Picture         =   "frmModInv.frx":1C9E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Visualizar precios de venta"
         Top             =   70
         Width           =   700
      End
      Begin VB.CommandButton CmdExporta 
         Caption         =   "Exportar"
         Height          =   480
         Left            =   4050
         Picture         =   "frmModInv.frx":1DA0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exportar existencias a archivo DBF"
         Top             =   70
         Width           =   700
      End
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "Actual"
         Height          =   480
         Left            =   2580
         Picture         =   "frmModInv.frx":1EE2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Actualizar datos para reflejar cambios realizados por otros usuarios"
         Top             =   80
         Width           =   700
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "&Sel.prov"
         Height          =   480
         Left            =   1890
         Picture         =   "frmModInv.frx":1FE4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Muestra inventario por proveedor"
         Top             =   70
         Width           =   700
      End
      Begin VB.CommandButton btninicializa 
         Caption         =   "&Ajustes"
         Height          =   480
         Left            =   1200
         Picture         =   "frmModInv.frx":2156
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ajustar inventario"
         Top             =   70
         Width           =   700
      End
      Begin VB.CheckBox chkcotiza 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cotiza Credito"
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
         Height          =   230
         Left            =   10080
         TabIndex        =   27
         Top             =   320
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmdUltimo 
         Height          =   300
         Left            =   120
         Picture         =   "frmModInv.frx":22D8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ir al último"
         Top             =   0
         Width           =   500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "No Fact"
         Height          =   480
         Left            =   3360
         Picture         =   "frmModInv.frx":244A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Productos pendientes de capturar"
         Top             =   70
         Width           =   700
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Productos XX"
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
         Left            =   10080
         TabIndex        =   16
         Top             =   30
         Width           =   1695
      End
   End
   Begin VB.Frame fratda 
      BackColor       =   &H80000001&
      Caption         =   "Seleccion de Inventario  de Tienda:"
      Enabled         =   0   'False
      Height          =   2415
      Left            =   4200
      TabIndex        =   28
      Top             =   2040
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cmbtiendas 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton btncambia 
         Caption         =   "Ok"
         Height          =   375
         Left            =   3960
         TabIndex        =   32
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox CHKTODAS 
         Caption         =   "TODAS LAS TIENDAS"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   2655
      End
      Begin VB.OptionButton opt 
         Caption         =   "Bodega"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         Caption         =   "Piso"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   29
         Top             =   1560
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdModInv 
      Bindings        =   "frmModInv.frx":278C
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      ForeColor       =   4194368
      HeadLines       =   1.5
      RowHeight       =   16
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "inprod"
         Caption         =   "  CLAVE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "BARRASPZA"
         Caption         =   "COD.BARRA PZA."
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
         DataField       =   "Descripc"
         Caption         =   "                              DESCRIPCION"
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
         DataField       =   "medida"
         Caption         =   "         MEDIDA"
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
      BeginProperty Column04 
         DataField       =   "InInicialP"
         Caption         =   "INV.INI.PZAS."
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
         DataField       =   "InCant"
         Caption         =   "EXIST. CAJAS"
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
         DataField       =   "InCantPza"
         Caption         =   "EXIST.PZAS."
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
         DataField       =   "incantcdc"
         Caption         =   "Cajas CDC"
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
      BeginProperty Column08 
         DataField       =   "incantpzaCdc"
         Caption         =   "Piezas CDC"
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
      BeginProperty Column09 
         DataField       =   "InInicial"
         Caption         =   "INV.INI. CAJAS"
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
      BeginProperty Column10 
         DataField       =   "minimo"
         Caption         =   "MINIMO"
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
      BeginProperty Column11 
         DataField       =   "maximo"
         Caption         =   "MAXIMO"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   4844.977
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   975.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoCotiza 
      Height          =   330
      Left            =   8880
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "AdoCotiza"
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
   Begin MSAdodcLib.Adodc Adodbf 
      Height          =   330
      Left            =   5280
      Top             =   5280
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
      Caption         =   "Adodbf"
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
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   8280
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDetAju 
      Height          =   330
      Left            =   2880
      Top             =   5280
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "AdoDetAju"
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
   Begin MSComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   15
      Top             =   8265
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Click en el encabezado de columna ordena los datos"
            TextSave        =   "Click en el encabezado de columna ordena los datos"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   " F4 = Agregar Cotizaciones"
            TextSave        =   " F4 = Agregar Cotizaciones"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "F5 = Cotizaciones"
            TextSave        =   "F5 = Cotizaciones"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "F6 - Inv. bodegas"
            TextSave        =   "F6 - Inv. bodegas"
            Object.ToolTipText     =   "Muestra los inventarios de las Bodegas de Mayoreo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "F7 -Historial"
            TextSave        =   "F7 -Historial"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "F8 - Pendientes"
            TextSave        =   "F8 - Pendientes"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "F9 - Totales"
            TextSave        =   "F9 - Totales"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoModInv 
      Height          =   330
      Left            =   480
      Top             =   5280
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "AdoModInv"
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
   Begin VB.Frame fracotiza 
      Height          =   8295
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton cmdcambia 
         Caption         =   "&Cambiar precios"
         Height          =   495
         Left            =   4440
         Picture         =   "frmModInv.frx":27A4
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Borra los productos de la tabla de cotizaciones"
         Top             =   7560
         Width           =   1335
      End
      Begin VB.PictureBox Rpt 
         Height          =   480
         Left            =   720
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   50
         Top             =   7200
         Width           =   1200
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "&Borrar cotizacion"
         Height          =   495
         Left            =   2640
         Picture         =   "frmModInv.frx":2916
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Borra los productos de la tabla de cotizaciones"
         Top             =   7560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCotRpte 
         Caption         =   "&Reporte"
         Height          =   495
         Left            =   6240
         Picture         =   "frmModInv.frx":2A88
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vista preeliminar de la cotizacion"
         Top             =   7560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCotReg 
         Caption         =   "R&egresar"
         Height          =   495
         Left            =   7800
         Picture         =   "frmModInv.frx":2FBA
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Regresar a visualizar inventario"
         Top             =   7560
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adoprov 
         Height          =   330
         Left            =   720
         Top             =   6480
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         Enabled         =   0
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmModInv.frx":312C
         Height          =   7215
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   12726
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
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
         Caption         =   "PRODUCTOS A INCLUIR EN LA COTIZACION"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "consec"
            Caption         =   "CLAVE"
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
            DataField       =   "descripc"
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
            DataField       =   "medida"
            Caption         =   "MEDIDA"
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
            DataField       =   "PRECIOBOD"
            Caption         =   "PRECIO CAJA"
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
            DataField       =   "preciobodp"
            Caption         =   "PRECIO PZA."
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
         BeginProperty Column05 
            DataField       =   "INCANT"
            Caption         =   "CAJAS"
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
            DataField       =   "INCANTP"
            Caption         =   "PIEZAS"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5025.26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2190.047
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   959.811
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmModInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nCantAnt
Private lCon
Private CadInv As String

Private Sub btncambia_Click()
On Error GoTo Error:
fratda.Enabled = False
fratda.Visible = False

If CHKTODAS.Value = 0 Then
    'VERIFICANDO QUE EXISTA LA TABLA
    Me.Caption = "INVENTARIO BODEGA    --" & cmbtiendas.Text
    CLAVEINVENTARIO = Mid(cmbtiendas.Text, 2, 2)
    If CLAVEINVENTARIO = 10 Or CLAVEINVENTARIO = 3 Then
       If opt(1).Value Then
           CAD = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIOpiso as inventario, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
       Else
           CAD = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
       End If
    Else
       If opt(1).Value Then
            CAD = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIOpiso   as inventario , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
       Else
           CAD = "SELECT TFPRODUC.activo, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.incant, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, incantcdc, incantpzacdc  FROM INVENTARIO" & Trim(CLAVEINVENTARIO) & "  as inventario , TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 ORDER BY DESCRIPC,CONTENID"
       End If
    End If
    'MsgBox cad
    AdoModInv.RecordSource = CAD
    AdoModInv.Refresh
    lblInfo.Caption = "Productos " + Str(AdoModInv.Recordset.RecordCount)
    Exit Sub
Else
    frminvtodo.Show
    Exit Sub
End If
Error:
     MsgBox Err.Description
     stb1.Panels(1).Text = "No se encontro el Inventario de esta Bodega ..."
     Unload Me
End Sub

Private Sub btninicializa_Click()
fraCon.Enabled = True
fraCon.Visible = True
txtContra.SetFocus
End Sub

Private Sub btnresp_Click()
If Me.txtContra.Text = "MOYLEON" Then
    cn.Execute "respaldainventario "
    MsgBox "Proceso de Respaldo Realizado ", vbInformation
Else
   MsgBox "No es posible Inicializar el inventario", vbCritical
End If
fraCon.Enabled = False
fraCon.Visible = False
txtContra.Text = ""
End Sub

Private Sub cmbproved_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub

Private Sub cmbProved_Validate(Cancel As Boolean)
Dim N As Integer
If cmbProved.Text = "_TODOS " Then
    todos = True
    txtclave.Text = ""
End If
 If todos = False Then
   If cmbProved.Text = "" Or IsNull(cmbProved.Text) Then
      MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
      cmbProved.SetFocus
      Cancel = True
   Else
      AdoProv.Recordset.MoveFirst
      AdoProv.Recordset.Find "NomProve = '" & cmbProved.Text & "'"
      If AdoProv.Recordset.EOF = True Then
         MsgBox "Debe seleccionar un proveedor de la lista desplegable", vbExclamation
      Else
        txtclave.Text = AdoProv.Recordset!prove
      End If
   End If
 Else
   Me.cmdAceptar.SetFocus
 End If

End Sub

Private Sub cmdAceptar_Click()
stb1.SimpleText = Space(40) + "Espere un momento obteniendo inventario de productos"
stb1.Refresh
If cmbProved.Text = "_TODOS " Or Trim(cmbProved.Text) = "" Then
    cCadena = "SELECT tfproduc.barraspza,incantpza,INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA,  INVENTARIO.insucursal, INVENTARIO.incant, INVENTARIO.inobserva, INVENTARIO.infeccaduprox, Inventario.Minimo, Inventario.Maximo FROM INVENTARIO, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND tfproduc.interno = 0 and tfproduc.activo = 1"
    todos = True
    AdoModInv.RecordSource = "SELECT INVENTARIO.incantcdc, INVENTARIO.incantpzacdc,tfproduc.barraspza,incantpza,INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA,  INVENTARIO.insucursal, INVENTARIO.incant, INVENTARIO.inobserva, INVENTARIO.infeccaduprox, Inventario.Minimo, Inventario.Maximo  FROM INVENTARIO, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND tfproduc.activo = 1 " & compInt & " ORDER BY descripc, contenid"
    Me.Caption = "INVENTARIO BODEGA:    " & cSucursal
Else
    Me.Caption = "INVENTARIO BODEGA:    " & cSucursal & Space(5) & "=> " & cmbProved.Text
    If Sql Then
       cCadena = "SELECT INVENTARIO.incantcdc, INVENTARIO.incantpzacdc,tfproduc.barraspza,incantpza,INVENTARIO.incantcdc, INVENTARIO.incantpzacdc, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA,  INVENTARIO.insucursal, INVENTARIO.incant, INVENTARIO.inobserva, INVENTARIO.infeccaduprox, Inventario.Minimo, Inventario.Maximo  FROM INVENTARIO, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND tfproduc.activo = 1 AND TfProduc.ClaProve = '" & txtclave.Text & "'"
       AdoModInv.RecordSource = "SELECT INVENTARIO.incantcdc, INVENTARIO.incantpzacdc,tfproduc.barraspza,incantpza,INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' +  lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA,  INVENTARIO.insucursal, INVENTARIO.incant, INVENTARIO.inobserva, INVENTARIO.infeccaduprox, Inventario.Minimo, Inventario.Maximo FROM INVENTARIO, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec " & compInt & " and tfproduc.activo = 1 AND TfProduc.ClaProve = '" & txtclave.Text & "' ORDER BY descripc,contenid"
       todos = False
    Else
       AdoModInv.RecordSource = "SELECT TFPRODUC.paquetes,INVENTARIO.incantcdc, INVENTARIO.incantpzacdc, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.InInicialP, INVENTARIO.incant, INVENTARIO.Ubicacion, INVENTARIO.inobserva, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, tfproduc.CLAPROVE  FROM INVENTARIO, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 AND TfProduc.ClaProve = '" & txtclave.Text & "' ORDER BY DESCRIPC,CONTENID"
    End If
End If
AdoModInv.CursorType = adOpenKeyset
AdoModInv.ConnectionString = cCadConex
AdoModInv.Refresh
lblInfo.Caption = "Productos " + Str(AdoModInv.Recordset.RecordCount)
Me.FraProv.Visible = False
Me.PicBotones.Visible = True
Me.dbgrdModInv.Visible = True
stb1.SimpleText = "                                           Click en el encabezado de columna ordena los datos en base a la columna seleccionada"

End Sub

Private Sub cmdBorrar_Click()
   If MsgBox("REALMENTE DESEAS ELIMINAR LA COTIZACION", vbQuestion + vbYesNo) = vbYes Then
      cn.Execute "DELETE FROM TPEDSUG"
      AdoCotiza.Refresh
   End If
End Sub

Private Sub cmdBuscaBarra_Click()
Dim cCve As String
Dim Antes
cCve = InputBox("Introduzca el codigo de barras a buscar", "Introducir codigo de barras")
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = dbgrdModInv.Bookmark
AdoModInv.Recordset.MoveFirst
AdoModInv.Recordset.Find "Barraspza =" & Trim(cCve)
If AdoModInv.Recordset.EOF Then
   MsgBox "El codigo de barras " & cCve & " no se encuentra en el inventario", vbExclamation
   dbgrdModInv.Bookmark = Antes
End If
dbgrdModInv.SetFocus

End Sub

Private Sub cmdBuscaCve_Click()
Dim cCve As String
Dim Antes
cCve = InputBox("Introduzca la clave a buscar", "Introducir clave")
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = dbgrdModInv.Bookmark
AdoModInv.Recordset.MoveFirst
AdoModInv.Recordset.Find "Inprod LIKE '" & Trim(cCve) & "*'"
If AdoModInv.Recordset.EOF Then
   MsgBox "La clave " & cCve & " no se encuentra en el inventario", vbExclamation
   dbgrdModInv.Bookmark = Antes
End If
dbgrdModInv.SetFocus
End Sub

Private Sub cmdBuscaDesc_Click()
Dim cCve As String
Dim Antes
cCve = InputBox("Introduzca la descripcion del producto a buscar", "Introducir descripcion")
If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
Antes = dbgrdModInv.Bookmark
AdoModInv.Recordset.MoveFirst
AdoModInv.Recordset.Find "DESCRIPC LIKE '" & Trim(cCve) & "*'"
If AdoModInv.Recordset.EOF Then
   MsgBox "La descripcion " & cCve & " no se encuentra en el inventario", vbExclamation
   dbgrdModInv.Bookmark = Antes
End If
dbgrdModInv.SetFocus
End Sub

Private Sub cmdcambia_Click()
RESP = InputBox("TECLEA LA ESCALA DE PRECIOS A COTIZAR" & Chr(13) & Chr(13) & "2.- Precio Mayoreo envío y/o crédito" & Chr(13) & "3.- Precio Mayoreo intermedio" & Chr(13) & "4.- Precio en bodega", "Escala", "2")
If Not IsNumeric(RESP) Then Exit Sub
cn.Execute "UPDATE tpedsug SET preciobod = precio" & Trim(RESP) & " FROM preprod where consec = preclave "
Me.AdoCotiza.Refresh
End Sub

Private Sub cmdCerrar_Click()
'Unload Me
Me.FraProv.Visible = False
End Sub

Private Sub cmdConAceptar_Click()
Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT * FROM USUARIOS WHERE PASSMASTER = '" & Trim(txtContra.Text) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
If RsCon.BOF And RsCon.EOF Then
   MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbExclamation
   txtContra.SetFocus
   SendKeys "+{HOME}"
   Exit Sub
Else
   If autoriza(RsCon!permisos, 1) Then
      frmModInvAjus.txtcampos(0).Text = RsCon!clave
      frmModInvAjus.Cmbusuario.Text = RsCon!Name
      lCon = True
      dbgrdModInv.Columns(5).Locked = False
      dbgrdModInv.Columns(6).Locked = False
      fraCon.Visible = False
   End If
End If
End Sub

Private Sub cmdConCance_Click()
  fraCon.Visible = False
End Sub

Private Sub cmdCotReg_Click()
  fracotiza.Visible = False
  Me.dbgrdModInv.SetFocus
End Sub

Private Sub cmdCotRpte_Click()
  Rpt.ReportFileName = App.Path & "\COTIZA.RPT"
  Rpt.Connect = cCadConex
  Rpt.WindowTitle = LCase("COTIZACION ELABORADA EL DIA " & UCase(Format(date, "LONG DATE")))
  Rpt.Formulas(0) = "ENCABEZADO = 'COTIZACION ELABORADA EL DIA " & UCase(Format(date, "LONG DATE")) & "'"
  Rpt.Action = 1
End Sub

Private Sub CmdExporta_Click()
'On Error GoTo Error:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cnFoxPro As ADODB.Connection
Dim lFranq As Boolean
   cMenAnt = stb1.Panels(1).Text
   Cmdlg.DialogTitle = "Seleccionar archivo para exportar existencias"
   Cmdlg.Filter = "Archivos Visual Fox pro (*.dbf) | *.dbf"
   Cmdlg.CancelError = True
   Cmdlg.ShowSave
   cRutArc = Cmdlg.FileName
   If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If
   stb1.Panels(1).Text = Space(45) & "Grabando archivo " & cRutArc
   stb1.Refresh
   
   For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
   Next
   cruta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.

   stb1.Panels(1).Text = "Limpiando archivo " & cArch
   stb1.Refresh
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set rsttemp = New ADODB.Recordset
   rsttemp.Open "SELECT claprove,consec,descripc,paquetes,contenid,medida,incant,incantpza,precio1,precio2,precio3,precio4 FROM Inventario, Tfproduc,Preprod WHERE Inprod = Consec AND preclave = consec AND preclave = inprod AND InCant > 0 AND INPROD > 0", cn, adOpenStatic, adLockOptimistic, adCmdText
   Set f = fs.GetFile("P:\ESTINVEN.DBF")
   f.Copy cRutArc, True
   AdoDbf.CommandType = adCmdText
   AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
   AdoDbf.RecordSource = "SELECT * FROM " & cArch
   AdoDbf.Refresh
   nreg = 1: NTOTREG = CStr(rsttemp.RecordCount)
   While Not rsttemp.EOF
      stb1.Panels(1).Text = "Exportando clave:" & CStr(rsttemp!CONSEC) & Space(5) & "Producto: " & CStr(nreg) & " de " & NTOTREG
      stb1.Refresh
      AdoDbf.Recordset.AddNew
      AdoDbf.Recordset!claprove = IIf(IsNull(rsttemp!claprove), "", Trim(rsttemp!claprove))
      AdoDbf.Recordset!CONSEC = Trim(rsttemp!CONSEC)
      AdoDbf.Recordset!descripc = rsttemp!descripc
      AdoDbf.Recordset!PAQUETES = IIf(IsNull(rsttemp!PAQUETES), 0, Trim(rsttemp!PAQUETES))
      AdoDbf.Recordset!CONTENID = IIf(IsNull(rsttemp!CONTENID), 0, Val(rsttemp!CONTENID))
      AdoDbf.Recordset!medida = IIf(IsNull(rsttemp!medida), "", Trim(rsttemp!medida))
      AdoDbf.Recordset!exicaja = Val(rsttemp!InCant)
      AdoDbf.Recordset!exipza = Val(rsttemp!InCantPza)
      AdoDbf.Recordset!precio1 = Val(rsttemp!precio1)
      AdoDbf.Recordset!PRECIO2 = Val(rsttemp!PRECIO2)
      AdoDbf.Recordset!PRECIO3 = Val(rsttemp!PRECIO3)
      AdoDbf.Recordset!precio4 = Val(rsttemp!precio4)
      AdoDbf.Recordset!pedirc = 0
      AdoDbf.Recordset!pedirp = 0
      AdoDbf.Recordset.Update
      rsttemp.MoveNext
      nreg = nreg + 1
   Wend
   AdoDbf.Recordset.Close
   Set AdoDbf.Recordset = Nothing
   Set f = fs.GetFile(cRutArc)
   f.Copy cruta & "\EXICAB.DBF", True
   stb1.Panels(1).Text = cMenAnt
   MsgBox "SE EXPORTARON " & NTOTREG & " PRODUCTOS", vbInformation
   Unload Me
  Exit Sub
Error:
  If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
     MsgBox Err.Description
  End If
   stb1.Panels(1).Text = cMenAnt
End Sub

Private Sub cmdInvIni_Click()
'strcveprod = dbgrdModInv.Columns(0).Text
strcveprod = AdoModInv.Recordset!Inprod
lpprov = True
lpprod = False
If InStr(1, UCase(computadora()), "MODVTA") = 0 Then
   frmprecios.Show
   'fnewprec.Show
Else
   fnewprec.Show
End If
End Sub

Private Sub cmdpenreg_Click()
  frapend.Visible = False
End Sub

Private Sub cmdprod_Click()
  lpprod = True
  strcveprod = AdoModInv.Recordset!Inprod
  frmprod.Show 1
End Sub

Private Sub CmdRefresh_Click()
Dim MenAnt
MenAnt = stb1.Panels(1).Text
stb1.SimpleText = Space(55) & "Espere un momento actualizando inventario..."
stb1.Refresh
AdoModInv.RecordSource = CadInv & ConInv & " ORDER BY descripc,contenid"
AdoModInv.Refresh
Me.lblInfo.Caption = "Productos: " & AdoModInv.Recordset.RecordCount
Me.lblInfo.Refresh
stb1.Panels(1).Text = MenAnt
stb1.Refresh
End Sub

Private Function ConInv() As String
If optVerInv(0).Value Then
   ConInv = " "
ElseIf optVerInv(1).Value Then
   ConInv = " AND TFPRODUC.activo = 1 "
ElseIf optVerInv(2).Value Then
   ConInv = " AND (INVENTARIO.INCANT > 0 OR INVENTARIO.INCANTPZA > 0) "
End If
End Function

Private Sub cmdRpteInv_Click()
On Error GoTo Error:
  mensa = stb1.Panels(1).Text
  stb1.Panels(1).Text = Space(55) & "Espere un momento generando reporte de existencias..........."
  stb1.Refresh
  Rpt.Connect = cCadConex
  Rpt.ReportFileName = App.Path & "\Exipro.rpt"
  If todos = True Then
     'EN EL CASO DE QUE SE TRATE DE TODOS LOS PRODUCTOS
      Rpt.WindowTitle = "EXISTENCIAS TOTALES "
      Rpt.Formulas(0) = "FORMSELEC = {TFPRODUC.CLAPROVE} <> '000' "
  Else
      Rpt.WindowTitle = "EXISTENCIAS DE " & cmbProved.Text
      Rpt.Formulas(0) = "FORMSELEC = {TFPRODUC.CLAPROVE} = '" & Me.txtclave.Text & "'"
  End If
  Rpt.Action = 1
  stb1.Panels(1).Text = mensa
  stb1.Refresh
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Command1_Click()
If Me.txtContra.Text = "MOYLEON" Then
    cn.Execute "INVENTARIOINICIAL '" & date & "'"
    MsgBox "Proceso Realizado", vbInformation
Else
   MsgBox "No es posible Inicializar el inventario", vbCritical
End If
fraCon.Enabled = False
fraCon.Visible = False
txtContra.Text = ""
End Sub

Private Sub Command2_Click()
Call muestrahistorico
End Sub


Private Sub Command3_Click()
frapend.Visible = True
AdoPend.ConnectionString = cCadConex
AdoPend.CommandType = adCmdText
AdoPend.RecordSource = "SELECT cnombre,d.noventa,v.fecha,folpreventa,d.cantidad,cantidadp,d.precio,d.importe FROM Ventas v, ventas_det d, catcliente WHERE agente = cclave and v.noventa =  d.noventa And D.facturado = 0 and  cl_producto =  '" & AdoModInv.Recordset!Inprod & "' and fecha >= '01/01/2003' Order by d.noventa"
AdoPend.Refresh
Me.DBGRDPEND.Caption = Me.AdoModInv.Recordset!descripc & "  " & AdoModInv.Recordset!medida
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
SendKeys "{DOWN}"
End Sub

Private Sub dbgrdModInv_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
  nCantAnt = AdoModInv.Recordset!InCant
End Sub

Private Sub dbgrdModInv_KeyDown(KeyCode As Integer, Shift As Integer)
Dim clave
Dim rs As ADODB.Recordset
On Error GoTo Error:
clave = frmModInv.AdoModInv.Recordset!Inprod
If KeyCode = 115 And stb1.Panels(2).Enabled Then        'Tecla de funcion F4
   If MsgBox("REALMENTE DESEAS ENVIAR A COTIZACIONES EL PRODUCTO" & Chr(13) & AdoModInv.Recordset!descripc & Space(5) & AdoModInv.Recordset!medida, vbQuestion + vbYesNo) = vbYes Then
      Set rs = New ADODB.Recordset
      rs.Open "SELECT LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, DESCRIPC, CONSEC, PRECIO2, PRECIO4,precio1,paquetes FROM PREPROD,TFPRODUC WHERE consec = preclave AND preclave = '" & AdoModInv.Recordset!Inprod & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
      nprecio = IIf(chkcotiza.Value = 1, rs!PRECIO2, rs!precio4)
      cn.Execute "INSERT INTO tpedsug(descripc,medida,consec,preciobod,preciobodp,INCANT,INCANTP,paquetes) VALUES ('" & rs!descripc & "','" & rs!medida & "','" & rs!CONSEC & "'," & nprecio & "," & rs!precio1 & ",0,0," & rs!PAQUETES & ")"
      Set rs = Nothing
      fracotiza.Visible = True
      AdoCotiza.ConnectionString = cCadConex
      AdoCotiza.CommandType = adCmdText
      AdoCotiza.RecordSource = "SELECT * FROM tpedsug order BY descripc"
      AdoCotiza.Refresh
   End If
ElseIf KeyCode = 116 And stb1.Panels(3).Enabled Then     'Tecla de funcion F5
   fracotiza.Visible = True
   AdoCotiza.ConnectionString = cCadConex
   AdoCotiza.CommandType = adCmdText
   AdoCotiza.RecordSource = "SELECT * FROM tpedsug order BY descripc"
   AdoCotiza.Refresh
ElseIf KeyCode = 117 Then 'F6 Muestra de Inventario de las Bodegas
   cmen = stb1.Panels(1).Text
   stb1.Panels(1).Text = "Obteniendo inventarios de Bodegas de Mayoreo"
   stb1.Refresh
   frminvtodo.Show
   stb1.Panels(1).Text = cmen
   stb1.Refresh
ElseIf KeyCode = 123 Then  'Tecla de funcion   F5
   Set rs = New ADODB.Recordset
   rs.Open "SELECT SUM(InCant * Precosto) as ImpInv, COUNT(consec) AS NumPro, SUM(InCant) NumCaj FROM Inventario,Tfproduc WHERE Inprod = consec AND inCant > 0 ", cn, adOpenKeyset, adLockOptimistic, adCmdText
   MsgBox "INFORMACION DEL INVENTARIO " & Chr(13) & Chr(13) & "PRODUCTOS   :   " & Format(rs!Numpro, "###,###,###") & Chr(13) & "CAJAS               :   " & Format(rs!NumCaj, "###,###,###.00") & Chr(13) & "IMPORTE          :   " & Format(rs!ImpInv, "$###,###,###,###.00"), vbInformation
ElseIf KeyCode = 118 Then  'Tecla de funcion   F7
    Call muestrahistorico
ElseIf KeyCode = 119 Then  'Tecla de funcion   F8
   Call muestrapendientes
ElseIf KeyCode = 120 Then  'Tecla de funcion   F9
   ' If CLAVEINVENTARIO = 10 Or CLAVEINVENTARIO = 3 Then
        CAD = " select sum(incant) as cajas, sum(incantpza) as piezas, count(consec) as variedad, sum(precosto * incant) as costocaj , sum((precosto / paquetes) * incantpza) as costopza , sum(precio1 * paquetes * incant) as precaja , sum(precio1 * incantpza ) as prepza from tfproduc, preprod, inventario where consec = preclave and consec = inprod and (incant >  0 or incantpza > 0 )"
   ' Else
    '    CAD = " select sum(incant) as cajas, sum(incantpza) as piezas, count(consec) as variedad, sum(precosto * incant) as costocaj , sum((precosto / paquetes) * incantpza) as costopza , sum(precio4 * incant) as precaja , sum(precio1 * incantpza ) as prepza from tfproduc, preprod, inventario" & Trim(CLAVEINVENTARIO) & "  as inventario where consec = preclave and consec = inprod and (incant >  0 or incantpza > 0 )"
    'End If
    Set rs = New ADODB.Recordset
    rs.Open CAD, cn, adOpenKeyset, adLockOptimistic, adCmdText
    CADENA = ""
    If Not rs.EOF Then
       CADENA = " " & vbCrLf
       CADENA = CADENA & " INVENTARIO DE BODEGA" & vbCrLf & vbCrLf
       CADENA = CADENA & " VARIEDADES  : " & vbTab & vbTab & rs!VARIEDAD & vbCrLf & vbCrLf
       CADENA = CADENA & " CAJAS       : " & vbTab & vbTab & rs!cajas & vbCrLf & vbCrLf
       CADENA = CADENA & " PIEZAS      : " & vbTab & vbTab & rs!piezas & vbCrLf & vbCrLf
       CADENA = CADENA & " COSTO       : " & vbTab & vbTab & Format(rs!costocaj + rs!COSToPZA, "$###,##0.00") & vbCrLf & vbCrLf
       CADENA = CADENA & " PRECIO VENTA: " & vbTab & vbTab & Format(rs!PRECAJA + rs!PREPZA, "$###,##0.00") & vbCrLf & vbCrLf
       CADENA = CADENA & vbCrLf & vbCrLf
    End If
    MsgBox CADENA, vbInformation, "TOTALES"
    rs.Close
End If
Exit Sub
Error:
    MsgBox Err.Description
End Sub

Private Sub muestrapendientes()
frmHisto.AdoEntPend.ConnectionString = cCadConex
   frmHisto.AdoEntPend.CommandType = adCmdText
   frmHisto.AdoEntPend.RecordSource = "SELECT * FROM Pedprove, detalleglobal WHERE dg_pedido = pp_pedido AND pp_recibe = 0 AND DG_PRODUCTO = '" & frmModInv.AdoModInv.Recordset!Inprod & " AND pp_sucursal = '" & Trim(Mid(cSucursal, 1, 3)) & "' ORDER BY pp_fecConfirma DESC"
   frmHisto.AdoEntPend.Refresh
   frmHisto.AdoEntSur.ConnectionString = cCadConex
   frmHisto.AdoEntSur.CommandType = adCmdText
   frmHisto.AdoEntSur.RecordSource = "SELECT * FROM Pedprove, detalleglobal WHERE dg_pedido = pp_pedido AND pp_recibe = 1 AND DG_PRODUCTO = '" & frmModInv.AdoModInv.Recordset!Inprod & "' AND dg_cantsol <> dg_cantreal AND pp_sucursal = '" & Trim(Mid(cSucursal, 1, 3)) & "' ORDER BY pp_fecrecibe DESC"
   frmHisto.AdoEntSur.Refresh
   frmHisto.DatGSal.Visible = False
   frmHisto.DatgEnt.Visible = False
   frmHisto.frapend.Visible = True
   frmHisto.Show 1
End Sub

Private Sub muestrahistorico()
On Error GoTo Error:
   stb1.Panels(1).Text = "Obteniendo Historial"
   stb1.Refresh
   clave = frmModInv.AdoModInv.Recordset!Inprod
   cn.Execute "DELETE FROM HistEnt"
   'Pedidos por proveedor
   If tipotienda <> 3 Then
      CAD = "INSERT INTO histent (folio, fechaelab, cantsol, cantrec, fechaconf,fecharec,facturas,importe,tipo,existencia) SELECT pp_pedido, pp_fechagen, dg_cantsol, dg_cantreal + dg_promocionr , pp_fecconfirma, pp_fecrecibe, factura1, impfac1, tipo = 'PEDPROVE', dg_existencia FROM Pedprove, detalleglobal, Notaentrada WHERE dg_pedido = pp_pedido AND pedido  = pp_pedido AND pp_recibe = 1 AND DG_CANTREAL > 0 AND DG_PRODUCTO = '" & clave & "' "
      cn.Execute CAD
   End If
   'Pedidos sugeridos instantaneos (PEDINST)
   CAD = "INSERT INTO histent (folio, fechaelab, cantsol, cantrec, fechaconf,fecharec,facturas,importe,tipo,existencia) SELECT p_pedido, P_FECPED, df_cantsol, df_cantreal, p_fecconfirma, p_fecentreal, factura1, impfac1, tipo = 'PEDINST',df_existencia FROM Pedidos, detalleFactura, notaentrada WHERE df_pedido = p_pedido AND pedido = p_pedido AND p_recibido = 1 AND df_CANTREAL > 0 AND df_prod = '" & clave & "' "
   cn.Execute CAD
   'Recibo de mercancia por traslados (TRASLREC)
   CAD = "INSERT INTO histent (folio, fechaelab, cantsol, cantrec, fechaconf,fecharec,facturas,importe,tipo) SELECT t_clave, null, 0, dt_cantidad, null, t_fecha, t_foliotie, t_costo, tipo = 'TRASLREC' FROM Traslados, DetalleTraslado WHERE t_clave = dt_clave AND t_motivocancela is null AND Dt_cantidad > 0 AND Dt_producto = '" & clave & "' AND t_entrada = 1 AND t_enviado = 1"
   cn.Execute CAD
   'Ajuste de inventario (AJUSTE)
   cn.Execute "INSERT INTO histent (folio, fecharec, cantrec,tipo,existencia) SELECT a_clave, a_fecha, da_cantidad, tipo = 'AJUSTE', da_cantidadant FROM Ajustes, DetalleAjustes WHERE a_clave = da_clave AND Da_producto = '" & clave & "' and da_cantidad > 0"
   'SALIDAS
   cn.Execute "INSERT INTO histent (folio, fecharec, cantrec,tipo,existencia,entrada) SELECT a_clave, a_fecha, da_cantidad, tipo = 'AJUSTE', da_cantidadant,0 FROM Ajustes, DetalleAjustes WHERE a_clave = da_clave AND Da_producto = '" & clave & "' and da_cantidad < 0"
   'de mercancia a traves de ventas
   cn.Execute "INSERT INTO histent (folio, fecharec, cantsol,cantrec,cantrecp,tipo,entrada) SELECT vta = 0, CONVERT(CHAR,FECHA,5), sum(cantidad),sum(cantidad), sum(cantidadp), tipo = 'VENTAS',0 FROM ventas v, ventas_det d WHERE v.fecha >= '01-01-2002' and v.noventa = d.noventa and cl_producto = '" & clave & "' AND d.cancelado = 0 GROUP BY CONVERT(CHAR,FECHA,5)"
   'Recibo de mercancia por traslados (TRASLSAL)
   cn.Execute "INSERT INTO histent (folio, fechaelab, cantsol, cantrec,cantrecp, fechaconf,fecharec,facturas,importe,sucursal,tipo,entrada) SELECT t_clave,t_fecha,dt_cantidad,dt_cantidad,dt_cantidadp,null,t_fecha,null, null, t_sucursalreceptor, tipo = 'TRASLSAL', ent = 0 FROM Traslados, DetalleTraslado, Cattienda WHERE t_clave = dt_clave AND T_sucursalReceptor = ticlave AND t_motivocancela is null AND (Dt_cantidad > 0 or Dt_cantidadP > 0) AND Dt_producto = '" & frmModInv.AdoModInv.Recordset!Inprod & "' AND t_entrada = 0 AND T_ENVIADO = 1   ORDER BY t_fecha DESC"
   frmHisto.Show 1
Exit Sub
Error:
  MsgBox Err.Description
  Exit Sub
End Sub

Private Sub dbgrdModInv_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 If dbgrdModInv.SelBookmarks.Count > 0 Then dbgrdModInv.SelBookmarks.Remove 0
 dbgrdModInv.SelBookmarks.Add dbgrdModInv.RowBookmark(Me.dbgrdModInv.Row)
End Sub

Private Sub Form_Activate()
 Unload frmAreaRecibo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then frmCalc.Show   'F8
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Set rsttempt = New ADODB.Recordset
rsttempt.Open "SELECT TICLAVE,TIDESCRIP FROM CATTIENDA order by tidescrip ", cn, adOpenStatic, adLockOptimistic, adCmdText
While Not rsttempt.EOF
    If Val(rsttempt!ticlave) < 10 Then
       cmbtiendas.AddItem "[0" & Trim(rsttempt!ticlave) & "]" & "  " & Trim(rsttempt!tidescrip) & "  "
    Else
       cmbtiendas.AddItem "[" & Trim(rsttempt!ticlave) & "]" & "  " & Trim(rsttempt!tidescrip) & "  "
    End If
    rsttempt.MoveNext
Wend
rsttempt.Close
'CLAVEINVENTARIO = Mid(cSucursal, 1, 3)
CLAVEINVENTARIO = 16
Me.Caption = "INVENTARIO BODEGA:    " & cSucursal
lDatAJu = "": lCon = False
AdoProv.CursorType = adOpenKeyset
AdoProv.ConnectionString = cCadConex
AdoProv.RecordSource = "select * from CATPROV"
AdoProv.Refresh
cmbProved.AddItem "_TODOS "
 Do While Not AdoProv.Recordset.EOF
    If Not IsNull(AdoProv.Recordset!NOMPROVE) Then
       cmbProved.AddItem AdoProv.Recordset!NOMPROVE
    End If
    AdoProv.Recordset.MoveNext
 Loop
 todos = False

 AdoModInv.CursorType = adOpenKeyset
 If Sql Then
    CadInv = "SELECT TFPRODUC.paquetes, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid,10,3)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.InInicialP, INVENTARIO.incant, INVENTARIO.Ubicacion, INVENTARIO.inobserva,INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, tfproduc.CLAPROVE,INVENTARIO.INCANTCDC,INVENTARIO.INCANTPZACDC FROM INVENTARIO, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 "
 Else
    CadInv = "SELECT TFPRODUC.paquetes, TFPRODUC.BarrasPza, INVENTARIO.inprod, TFPRODUC.Descripc,  LTRIM(STR(TFPRODUC.paquetes)) + ' X ' + lTrim(str(TFPRODUC.contenid)) + space(2) + TFPRODUC.medida  AS MEDIDA, INVENTARIO.InInicial, INVENTARIO.InInicialP, INVENTARIO.incant, INVENTARIO.Ubicacion, INVENTARIO.inobserva, INVENTARIO.minimo, INVENTARIO.maximo,INVENTARIO.InCantPza, tfproduc.CLAPROVE  FROM INVENTARIO, TFPRODUC WHERE INVENTARIO.inprod = TFPRODUC.consec AND TFPRODUC.Paquetes > 0 "
 End If
 AdoModInv.RecordSource = CadInv & ConInv & " ORDER BY descripc, contenid"
 AdoModInv.ConnectionString = cCadConex
 AdoModInv.Refresh
 lblInfo.Caption = "Productos: " + Trim(Str(AdoModInv.Recordset.RecordCount))
End Sub

Private Sub CmdActualizar_Click()
Me.FraProv.Visible = True
txtclave.SetFocus
End Sub

Private Sub cmdAnterior_Click()
AdoModInv.Recordset.MovePrevious
If AdoModInv.Recordset.BOF Then AdoModInv.Recordset.MoveFirst
End Sub

Private Sub cmdPrimero_Click()
  AdoModInv.Recordset.MoveFirst
End Sub

Private Sub cmdRegresar_Click()
  Unload Me
End Sub

Private Sub cmdSiguiente_Click()
AdoModInv.Recordset.MoveNext
If AdoModInv.Recordset.EOF Then AdoModInv.Recordset.MoveLast
End Sub

Private Sub cmdUltimo_Click()
AdoModInv.Recordset.MoveLast
End Sub

Private Sub dbgrdModInv_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo Error:
nCantAnt = OldValue
If Not IsNumeric(lDatAJu) Then
   If Not IsNumeric(frmModInvAjus.txtcampos(0).Text) Then frmModInvAjus.txtcampos(0).Text = Mid(cCveDesUsu, 1, 3)
   frmModInvAjus.Show 1     'Si NO capturaron los datos correspondientes al ajuste
End If
If Not IsNumeric(lDatAJu) Or ((InStr(1, "ABARROTES 1-ABARROTES 2-ABARROTES 3-ARTICULOS DE LIMPIEZA-DESPENSA", Trim(Mid(AdoModInv.Recordset!descripc, 1, 15)))) And InStr(1, "19-54-5", Trim(lDatAJu)) = 0) And Me.dbgrdModInv.Columns(ColIndex).DataField = "INCANT" Then
   MsgBox "Al NO guardar los datos del ajuste no puede modificar el inventario, O NO  esta autorizado para modificar este producto", vbExclamation
   Cancel = True
Else
   If UCase(dbgrdModInv.Columns(ColIndex).DataField) = "INCANT" Then   'Columna correspondiente a cantidad por caja
      CmdActualizar.Enabled = True
      AdoDetAju.Recordset.AddNew
      AdoDetAju.Recordset!da_clave = nOp
      AdoDetAju.Recordset!da_producto = AdoModInv.Recordset!Inprod
      AdoDetAju.Recordset!da_cantidadAnt = nCantAnt
      AdoDetAju.Recordset!da_cantidad = dbgrdModInv.Columns(ColIndex).Text - AdoModInv.Recordset!InCant
      AdoDetAju.Recordset.Update
   End If
End If
Exit Sub
Error:
  AdoDetAju.Refresh
  MsgBox Err.Description
End Sub

Private Sub dbgrdModInv_HeadClick(ByVal ColIndex As Integer)
Dim cCampo As String
     stb1.Panels(1).Text = Space(30) & "Espere un momento ordenando datos por la columna  " & Trim(dbgrdModInv.Columns(ColIndex).Caption)
     AdoModInv.RecordSource = CadInv & ConInv & " ORDER BY " & Trim(dbgrdModInv.Columns(ColIndex).DataField)
     AdoModInv.Refresh
     stb1.Panels(1).Text = Space(70) & "Datos ordenados por la columna " & Trim(dbgrdModInv.Columns(ColIndex).Caption)
     lblInfo.Caption = "Productos: " + Trim(Str(AdoModInv.Recordset.RecordCount))
     SendKeys "{ENTER}"
End Sub

Private Sub Form_Resize()
On Error GoTo Error:
 dbgrdModInv.Width = frmModInv.ScaleWidth - 400
 dbgrdModInv.Height = frmModInv.ScaleHeight - 1100
Error:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  frmAreaRecibo.Show
End Sub


Private Sub txtclave_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  cmdCerrar_Click
ElseIf KeyAscii = 13 Then
  KeyAscii = 0
  SendKeys vbTab
End If

End Sub

Private Sub txtclave_LostFocus()
On Error Resume Next
 txtclave.Text = Trim(UCase(txtclave.Text))
 txtclave.Refresh
 If txtclave.Text = "" Or IsNull(txtclave.Text) Then
    cmbProved.SetFocus
    Exit Sub
 Else
    AdoProv.Recordset.MoveFirst
    AdoProv.Recordset.Find " Prove = '" & Trim(txtclave.Text) & "'"
    If AdoProv.Recordset.EOF = True Then
       'MsgBox "No existe la clave del proveedor especificado", vbExclamation
       cmbProved.SetFocus
       Exit Sub
    End If
 End If
 cmbProved.Text = AdoProv.Recordset!NOMPROVE
 cmdAceptar.SetFocus
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    cmdConAceptar_Click
 ElseIf KeyAscii = 27 Then
    cmdConCance_Click
 End If
End Sub
