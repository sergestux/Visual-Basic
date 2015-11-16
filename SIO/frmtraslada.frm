VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtraslados 
   Caption         =   "Entradas y Salidas de productos"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "frmtraslada.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fratda 
      BackColor       =   &H80000001&
      Caption         =   "Seleccion de Tienda"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   3120
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton btntda 
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   3720
         TabIndex        =   36
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbFiltrotda 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   600
         Width           =   2895
      End
   End
   Begin MSAdodcLib.Adodc Adodbf 
      Height          =   330
      Left            =   0
      Top             =   2520
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
   Begin ComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   33
      Top             =   7980
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
   Begin VB.CommandButton cmdOpcion 
      Height          =   400
      Index           =   7
      Left            =   2760
      Picture         =   "frmtraslada.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Generar devolucion de mercancia a proveedores"
      Top             =   240
      Width           =   500
   End
   Begin VB.Frame fraCon 
      BackColor       =   &H00808000&
      Caption         =   "PROPORCIONE CONTRASEÑA"
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   3960
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblContra 
         BackColor       =   &H00808000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc AdoTraslados 
      Height          =   375
      Left            =   480
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoTraslados"
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
   Begin VB.Frame framenu 
      Height          =   732
      Left            =   240
      TabIndex        =   15
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton CmdMoverse 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   1440
         Picture         =   "frmtraslada.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Ultimo"
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   0
         Picture         =   "frmtraslada.frx":16EE
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Siguiente"
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   720
         Picture         =   "frmtraslada.frx":1860
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Anterior"
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3120
         Picture         =   "frmtraslada.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Primero"
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   5
         Left            =   4920
         Picture         =   "frmtraslada.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Presentacion preelimar de los traslados seleccionados en el rango especificado"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Height          =   400
         Index           =   4
         Left            =   4320
         Picture         =   "frmtraslada.frx":2076
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Buscar clave del traslado en el rango seleccionado"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton CmdMoverse 
         Enabled         =   0   'False
         Height          =   400
         Index           =   6
         Left            =   4080
         Picture         =   "frmtraslada.frx":2170
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Actualiza precios del traslado seleccionado"
         Top             =   0
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton CmdRefresh 
         Height          =   400
         Left            =   5520
         Picture         =   "frmtraslada.frx":2272
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Actualiza pantalla de traslados"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdimporta 
         Height          =   400
         Left            =   6120
         Picture         =   "frmtraslada.frx":2374
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Importar el traslados seleccionado"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdexport 
         Enabled         =   0   'False
         Height          =   400
         Left            =   6720
         Picture         =   "frmtraslada.frx":24B6
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Enviar Traslados a Tiendas"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton btntda1 
         Height          =   400
         Left            =   7320
         Picture         =   "frmtraslada.frx":25F8
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Selecciona Envios de Una Tienda"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton btnvta 
         Height          =   400
         Left            =   7920
         Picture         =   "frmtraslada.frx":2A3A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Genera Facturas Con Base a Traslados"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   3
         Left            =   10920
         Picture         =   "frmtraslada.frx":2E7C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Cancelar todo el enio y aumentar Inventario"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   8
         Left            =   3120
         Picture         =   "frmtraslada.frx":3186
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modificar devolucion a proveedores"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   6
         Left            =   1920
         Picture         =   "frmtraslada.frx":3490
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Modificar recepcion de traslado"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Mod. Cant."
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   7
         ToolTipText     =   "Modificar cantidades del traslado conopcion de poner en cero el inventario"
         Top             =   120
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   4
         Left            =   3720
         Picture         =   "frmtraslada.frx":379A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancelar todo el enio y aumentar Inventario"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   1
         Left            =   720
         Picture         =   "frmtraslada.frx":3AA4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Modificar envio capturado"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   0
         Left            =   120
         Picture         =   "frmtraslada.frx":3C16
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Capturar un nuevo envio"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   400
         Index           =   2
         Left            =   1320
         Picture         =   "frmtraslada.frx":3D58
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Recibir traslado de productos"
         Top             =   240
         Width           =   500
      End
      Begin VB.ComboBox cmbFiltro 
         Height          =   315
         ItemData        =   "frmtraslada.frx":4062
         Left            =   8520
         List            =   "frmtraslada.frx":4064
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblFiltro 
         Alignment       =   2  'Center
         Caption         =   "Filtro"
         Height          =   255
         Left            =   8400
         TabIndex        =   16
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Frame fradescripcion 
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   11535
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   0
         Left            =   8040
         TabIndex        =   51
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   65273859
         CurrentDate     =   37257
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Index           =   1
         Left            =   9600
         TabIndex        =   52
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   65273859
         CurrentDate     =   37257
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de Traslados: 999,999"
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
         Left            =   7080
         TabIndex        =   46
         Top             =   120
         Width           =   735
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Fecha Fin. Env."
         Height          =   255
         Index           =   1
         Left            =   9720
         TabIndex        =   32
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         Caption         =   "Fecha Ini. Env."
         Height          =   255
         Index           =   0
         Left            =   7920
         TabIndex        =   31
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblSucE 
         BackColor       =   &H80000004&
         Caption         =   "Sucursal Emisora    :"
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
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblSucEmi 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblSucR 
         BackColor       =   &H80000004&
         Caption         =   "Sucursal Receptora:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblSucrec 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   4335
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame FraCancelar 
      Caption         =   "Cancelar etiquetas"
      Height          =   3495
      Left            =   2640
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox txtCampos 
         Height          =   375
         Index           =   6
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2880
         Width           =   5535
      End
      Begin VB.OptionButton optopc 
         Caption         =   "Cancelar &salida (caja)"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton optopc 
         Caption         =   "Cancelar &entrada (caja)"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   20
         Top             =   1080
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   375
         Left            =   5640
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtCancela 
         Height          =   315
         Left            =   3360
         TabIndex        =   18
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label LBLETIQUETAS 
         Caption         =   "Motivo"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   25
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label LBLETIQUETAS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   720
         TabIndex        =   24
         Top             =   1920
         Width           =   5895
      End
      Begin VB.Label LBLETIQUETAS 
         Caption         =   "Codigo de barras de la caja:"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   23
         Top             =   480
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid dbgrdTras 
      Bindings        =   "frmtraslada.frx":4066
      Height          =   5865
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10345
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1.5
      RowHeight       =   15
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "t_clave"
         Caption         =   "FOL.UNI."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "t_foliotie"
         Caption         =   "FOL.TIE."
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
      BeginProperty Column02 
         DataField       =   "t_sucursalEmisor"
         Caption         =   "SUC. EMI."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "m/d/yy h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "t_fecha"
         Caption         =   "        FECHA DE ENVIO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yyyy hh:mm AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "t_tipo"
         Caption         =   "TRASL. ABTO."
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
      BeginProperty Column05 
         DataField       =   "t_monto"
         Caption         =   "IMPORTE"
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
         DataField       =   "t_costo"
         Caption         =   "     COSTO"
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
      BeginProperty Column07 
         DataField       =   "t_sucursalreceptor"
         Caption         =   "SUC.REC."
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
      BeginProperty Column08 
         DataField       =   "t_fecharec"
         Caption         =   "FECHA REC."
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
      BeginProperty Column09 
         DataField       =   "t_enviado"
         Caption         =   "ENV."
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
      BeginProperty Column10 
         DataField       =   "t_entrada"
         Caption         =   "ENTRAD"
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
      BeginProperty Column11 
         DataField       =   "t_venta"
         Caption         =   "Facturado"
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
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
   Begin VB.Menu x 
      Caption         =   "Traslados"
      WindowList      =   -1  'True
      Begin VB.Menu abrirar 
         Caption         =   "Abrir Archivo"
      End
      Begin VB.Menu agregart 
         Caption         =   "Agregar Traslado"
      End
      Begin VB.Menu eliminart 
         Caption         =   "Eliminar Traslado"
         HelpContextID   =   1
      End
      Begin VB.Menu cerrara 
         Caption         =   "Cerrar Archivo"
      End
      Begin VB.Menu fe 
         Caption         =   "Fecha"
      End
      Begin VB.Menu mnufacturar 
         Caption         =   "Facturar"
      End
      Begin VB.Menu qtafac 
         Caption         =   "Quitar Factura"
      End
   End
End
Attribute VB_Name = "frmtraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rstSuc As ADODB.Recordset
Private cCond As String    'Filtro del Control de datos al que esta enlazado el dbgrid
Private ccondrpt As String 'Filtro del Rpt
Private ntext  As Integer  'Para saber a que textbox se pasa al darle doble click al calendario
Private cFecha As String   'Filtro de rango de traslados por fecha
Private cFecharpt As String
Private nCancela As Integer
Private cfolCan As String
Private rfc As String

Private Sub abrirar_Click()
   Call cmdexport_Click
End Sub

Private Sub AdoTraslados_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF Then Exit Sub
rstSuc.MoveFirst
rstSuc.Find "ticlave = '" & AdoTraslados.Recordset!t_sucursalemisor & "'"
lblSucEmi.Caption = IIf(rstSuc.EOF, "", rstSuc!tidescrip)
rstSuc.MoveFirst
rstSuc.Find "ticlave = '" & AdoTraslados.Recordset!t_sucursalreceptor & "'"
lblSucrec.Caption = IIf(rstSuc.EOF, "", rstSuc!tidescrip)
End Sub

Private Sub agregart_Click()
Call exporta2
End Sub

Private Sub btntda_Click()
On Error GoTo Error:
 Select Case cmbFiltrotda.ListIndex
 Case 4  'todos excepto los cancelados
     cCond = "( t_sucursalreceptor = '13' or t_sucursalreceptor = '14' or t_sucursalreceptor = '15') AND t_enviado = 1 AND t_entrada = 0 AND t_motivocancela is null"
 Case 0  'Atzompa
     'cCond = "T_sucursalreceptor = 13 AND t_enviado = 1 AND t_entrada = 0 AND t_motivocancela is null"
     cCond = "T_sucursalreceptor = 13 AND t_enviado = 1 AND t_motivocancela is null"
 Case 1  'Miahuatlan
     'cCond = "T_sucursalreceptor = 15 AND t_enviado = 1 AND t_entrada = 0 AND t_motivocancela is null"
     cCond = "T_sucursalreceptor = 15 AND t_enviado = 1 AND t_motivocancela is null"
 Case 2  'Tlacolula
     'cCond = "T_sucursalreceptor = 14 AND t_enviado = 1 AND t_entrada = 0 AND t_motivocancela is null"
     cCond = "T_sucursalreceptor = 14 AND t_enviado = 1 AND t_motivocancela is null"
 Case 3  'Zimatlan
     'cCond = "T_sucursalreceptor = 27 AND t_enviado = 1 AND t_entrada = 0 AND t_motivocancela is null"
     cCond = "T_sucursalreceptor = 27 AND t_enviado = 1 AND t_motivocancela is null"
 End Select
' cCond = cCond
 cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
 
 If Sql Then
     cFecha = " AND (month(T_fecha) >= " & Month(dtpFecha(0).Value) & " and (day(T_fecha) > = " & Day(dtpFecha(0).Value) & cOper & " (day(T_fecha)<= " & Day(dtpFecha(1).Value) & " and month(T_fecha)<= " & Month(dtpFecha(1).Value) & ")) and year(T_fecha)>= " & Year(dtpFecha(0).Value) & " and year(T_fecha)<= " & Year(dtpFecha(1).Value) & ")"
  Else
    'COMPATIBILIDAD CON ACCESS
     Dim INI As Date
     Dim FIN As Date
     INI = dtpFecha(0).Value
     FIN = dtpFecha(1).Value
     'PRIMERO MES , DIA , AÑO
    INI = INI - 1
    FIN = FIN + 1
    'cFecha = " AND (month(T_fecha) >= " & Month(dtpfecha(0).value) & " and (day(T_fecha) > = " & Day(dtpfecha(0).value) & cOper & " (day(T_fecha)<= " & Day(dtpfecha(1).value) & " and month(T_fecha)<= " & Month(dtpfecha(1).value) & ")) and year(T_fecha)>= " & Year(dtpfecha(0).value) & " and year(T_fecha)<= " & Year(dtpfecha(1).value) & ")"
    cFecha = " AND  (((TRASLADOS.T_fecha) > #" & Format(INI, "mm /dd/yy") & "#)) AND  (((traslados.t_fecha)<#" & Format(FIN, "mm/dd/yy") & "#))"
     'FIN DE COMPATIBILIDAD
  End If
 CAD = "SELECT * FROM [TRASLADOS] WHERE " & cCond & cFecha
 'MsgBox cad
 AdoTraslados.RecordSource = CAD
 AdoTraslados.Refresh
' For n = 0 To 5
'   CmdMoverse(n).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
' Next
 cmdopcion(0).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 lblInfo.Caption = Str(AdoTraslados.Recordset.RecordCount)
 fratda.Enabled = False
 fratda.Visible = False
 Exit Sub
Error:
 MsgBox Err.Description
 
End Sub

Private Sub btntda1_Click()
Call cambiatienda
End Sub

Private Function agregafactura(fecha As Date, SERIE As String, Factura As Integer, venta As Integer, CLIENTE As String) As Boolean
Set rs = New ADODB.Recordset
rs.Open "select * from catcliente where cclave = " & CLIENTE, cn, adOpenDynamic, adLockOptimistic, adCmdText
rfc = rs!crfc
Set RSTEMP = New ADODB.Recordset
CAD = "select * from facventa where numfactura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
RSTEMP.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
total = 0: iva = 0: ieps = 0
If RSTEMP.EOF Then
    CAD = "insert into facventa(rfc,noventa,facfecha,faccliente,total,iva,ieps,numfactura,serie,cobrado,totfac,faccobro) values(" & _
           "'" & rfc & "'," & venta & ",'" & fecha & "','" & CLIENTE & "'," & total & "," & iva & "," & ieps & ",'" & Factura & "','" & SERIE & "',1," & total & ",'" & fecha & "')"
    'MsgBox CAD
    cn.Execute CAD
    'factura = ac + 1
    agregafactura = True
Else
    agregafactura = False
End If
End Function

Private Sub generafactura()
Dim venta As Integer
Set RSTEMP = New ADODB.Recordset
CAD = "select min(noventa) as venta from ventas "
RSTEMP.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
venta = RSTEMP!venta
Set RSTEMP = Nothing
Dim numfac As Integer
Dim Factura As String
Dim SERIE As String
Dim iva As Integer
Dim ieps As Integer
Dim CLIENTE As String
Open App.Path & "\envfac.TXT" For Input As #1
resp1 = MsgBox("Se ha detectado un archivo especial de traslados a facturar" & vbCrLf & "Desea Tomar en cuenta este archivo para Facturar", vbYesNo)
numfac = 0
'numprod = InputBox("Escriba el numero de Productos por Factura", "PRODUCTOS")
numprod = 15
numfac = InputBox("Escriba el numero de Factura Inicial...", "FACTURA INICIAL")
SERIE = InputBox("Escriba la serie ", "SERIE")
RESP = MsgBox("Deseas Generar las Facturas Con la fecha del Envio" & vbCrLf & "[SI= FECHA DE ENVIO]" & vbCrLf & "[NO=FECHA ACTUAL]", vbYesNo)
If IsNull(numfac) Or Val(numfac) < 1 Then
   MsgBox "Debe Especificar Un numero valido para Productos en Facturas", vbInformation
   Exit Sub
End If
If IsNull(numprod) Or Val(numprod) < 1 Then
   MsgBox "Debe Especificar Un numero valido para Productos en Facturas", vbInformation
   Exit Sub
End If
If IsNull(SERIE) Then
   MsgBox "Debe Especificar Una serie  valida", vbInformation
   Exit Sub
End If
Select Case Trim(AdoTraslados.Recordset!t_sucursalreceptor)
        Case "13"
             CLIENTE = "1"
        Case "14"
             CLIENTE = "4"
        Case "15"
             CLIENTE = "3"
        Case "27"
             CLIENTE = "5"
End Select
'se van a pedir factura por factura conforme se vayan ocupando
'el proceso agrupa los traslados seleccionados las facturas que quepan
'pero puede conbinar una o mas facturas de acuerdo al numero
'de productos que se hayan especificado
'se pone un campo temporal en t_venta se pone un numero 777
'para tomar en el select todos los productos
Dim fecha As Date
If RESP = vbYes Then
   fecha = dtpFecha(1).Value
Else
   fecha = date
End If
If Not agregafactura(fecha, SERIE, numfac, venta, CLIENTE) Then
        MsgBox "El numero de Factura " & numfac & " Ya existe, Corriga el dato Y vuelva a Procesar ", vbInformation
        Exit Sub
End If
Set rs = New ADODB.Recordset
cCond = " "
If resp1 = vbYes Then
    While Not EOF(1)
    Line Input #1, trasladox
    cCond = cCond & " t_clave = '" & Trim(trasladox) & "' or "
    Wend
    'le quitamos el ultimo and
    cCond = Mid(cCond, 1, Len(cCond) - 4)
    CAD = "SELECT * FROM Traslados WHERE " & cCond
    AdoTraslados.RecordSource = CAD
    AdoTraslados.Refresh
End If
'PRIMER PROCESO SE AGRUPAN LOS TRASLADOS SELECCIONADOS
Me.AdoTraslados.Recordset.MoveFirst
While Not AdoTraslados.Recordset.EOF
   If AdoTraslados.Recordset!t_venta = 0 Then
        AdoTraslados.Recordset!t_venta = -100
        AdoTraslados.Recordset.Update
   End If
   AdoTraslados.Recordset.MoveNext
Wend
'AHORA SE AGRUPAN LOS PRODUCTOS DONDE  T_VENTA = -100
'AdoTraslados.Recordset.Close
CAD = "SELECT max(dt_producto) producto, sum(dt_cantidad) cantidad ,sum(dt_cantidadp) cantidadp , max(dt_costo) costo, max(dt_costop) costop, max(dt_importe) importe, max(dt_venta) venta, max(dt_ventap) ventap, max(dt_tasaieps) tasaieps, max(dt_iva) iva, max(dt_ieps)ieps " & _
      "from detalletraslado, traslados where dt_clave = t_clave and t_venta = -100  and ( dt_cantidad > 0 or dt_cantidadp > 0 ) group by dt_producto "
rs.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
productos = 0
While Not rs.EOF
    productos = productos + 1
    If productos > Val(numprod) Then
        'Call ACTUALIZAMONTO(serie, factura)
        numfac = numfac + 1
        t = agregafactura(fecha, SERIE, numfac, venta, CLIENTE)
        productos = 1
    End If
    
    If IsNull(rs!ieps) Or rs!ieps < 0 Then
        ieps = 0
    Else
        ieps = rs!ieps
    End If
    If IsNull(rs!iva) Or rs!iva < 0 Then
        iva = 0
    Else
        iva = rs!iva
    End If
    x = rs!venta
    If rs!cantidad > 0 Or rs!cantidadp > 0 Then
            'tasaieps = validatasa(ieps, iva)
            tasaieps = rs!tasaieps
            importe = (rs!venta * rs!cantidad) + (rs!ventaP * rs!cantidadp)
            'MsgBox ""
            CAD = "INSERT INTO facventa_det(producto,cantidad,cantidadp,precio,preciop,costo,costop,importe,iva,ieps,tasaieps,factura,serie,venta,fecha_det,rfc_det)" & _
                  " values ( " & "'" & Trim(rs!producto) & "'," & rs!cantidad & "," & rs!cantidadp & "," & rs!venta & "," & rs!ventaP & "," & rs!costo & "," & rs!costop & "," & importe & "," & iva & "," & ieps & "," & tasaieps & ",'" & Trim(numfac) & "','" & Trim(SERIE) & "'," & venta & ",'" & fecha & "','" & rfc & "')"
            cn.Execute CAD
    End If
    rs.MoveNext
Wend
'LA ULTIMA FACTURA
'Call ACTUALIZAMONTO(serie, factura)
'SE PONEN LOS -100 COMO GENERADOS
cn.Execute "UPDATE TRASLADOS SET T_VENTA = " & numfac & " WHERE T_VENTA = -100 "
End Sub

Private Sub ACTUALIZAMONTO(SERIE As String, Factura As String)
'SE ACTUALIZAN LAS FACTURAS
CAD = " update facventa set iva = " & _
      " (select sum(importe - importe / (1 + (iva/100)) ) " & _
      " from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' ) " & _
      " Where numfactura = '" & Trim(Factura) & "' And serie = '" & Trim(SERIE) & "'"
cn.Execute CAD
'el ieps
CAD = "update facventa set ieps = " & _
      " (select sum( IMPORTE - (importe / (1 + (iva/100)) / (1 + (IEPS/100)) " & _
      " + (importe - importe / (1 + (iva/100)))  )  ) " & _
      " from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' )" & _
     " Where numfactura = '" & Trim(Factura) & "' And serie = '" & Trim(SERIE) & "'"
cn.Execute CAD
'SE ACTUALIZA EL MONTO DE LA FACTURA EN EL GLOBAL
CAD = "update facventa set total = (select sum(importe) from facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'  AND ( CANTIDAD > 0 OR CANTIDADP > 0 ) )  where numfactura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
cn.Execute CAD
End Sub

Private Sub btnvta_Click()
Close #1
generafactura
MsgBox "Proceso Finalizado...", vbInformation
Exit Sub

Set RSTEMP = New ADODB.Recordset
Set rs = New ADODB.Recordset
Dim iva As Integer
Dim ieps As Integer
MsgBox " ESTE PROCESO SOLO TOMA EN CUENTA LOS TRASLADOS CONFIRMADOS COMO ENVIADOS ", vbInformation, "GENERACION DE VENTA"
AdoTraslados.RecordSource = "SELECT * FROM Traslados WHERE " & cCond & cFecha & " ORDER BY T_FECHA"
AdoTraslados.Refresh
Me.dbgrdTras.Refresh
    cresp = MsgBox("DESEAS GENERAR LA VENTA CON FECHA DEL ENVIO " & Chr(13) & "[SI] = Genera venta con fecha de envio" & Chr(13) & "[NO] = Genera venta con fecha actual del sistema", vbYesNoCancel + vbQuestion)
    If cresp = vbCancel Then
       Exit Sub
    End If
    cn.BeginTrans
    AdoTraslados.Recordset.MoveFirst
    folVta = 0: fecha = "": MenAnt = Stb1.SimpleText
    While Not AdoTraslados.Recordset.EOF
    If AdoTraslados.Recordset!t_enviado Then '= 1 Then
      If AdoTraslados.Recordset!t_venta = 0 Then
         If fecha <> CStr(Day(AdoTraslados.Recordset!T_FECHA)) & CStr(Month(AdoTraslados.Recordset!T_FECHA)) & CStr(Year(AdoTraslados.Recordset!T_FECHA)) & Trim(AdoTraslados.Recordset!t_sucursalreceptor) Then
          Select Case Trim(AdoTraslados.Recordset!t_sucursalreceptor)
             Case "13"
                  CLIENTE = "1"
             Case "14"
                  CLIENTE = "4"
             Case "15"
                  CLIENTE = "3"
             Case "27"
                  CLIENTE = "5"
          End Select
                    'MsgBox folVta
                    'RSTEMP.Open "SELECT SUM(IMPORTE) AS importe FROM ventas_det WHERE NOVENTA =  " & folVta, cn, adOpenDynamic, adLockOptimistic, adCmdText
                    'cn.Execute "UPDATE VENTAS SET montototal = " & IIf(IsNull(RSTEMP!importe), 0, RSTEMP!importe) & " WHERE noventa = " & folVta
                    'RSTEMP.Close
                    If cresp = vbYes Then
                        cn.Execute "INSERT INTO ventas(fecha,clcliente,tipoventa,situacion,credito,facrfc,cancelado) VALUES('" & AdoTraslados.Recordset!T_FECHA & "','" & CLIENTE & "',1,'1',0,1,0)"
                    Else
                        cn.Execute "INSERT INTO ventas(fecha,clcliente,tipoventa,situacion,credito,facrfc,cancelado) VALUES('" & date + Time & "','" & CLIENTE & "',1,'1',0,1,0)"
                    End If
                    RSTEMP.Open "SELECT MAX(NOVENTA) as NOVENTA FROM VENTAS", cn, adOpenDynamic, adLockOptimistic, adCmdText
                    folVta = RSTEMP!noventa
                    RSTEMP.Close
       End If
       cn.Execute "UPDATE TRASLADOS SET t_venta = " & folVta & " WHERE T_CLAVE = '" & AdoTraslados.Recordset!t_clave & "'"
       rs.Open "SELECT * FROM detalletraslado WHERE dt_clave = '" & AdoTraslados.Recordset!t_clave & "' AND (dt_cantidad > 0 or dt_cantidadp > 0)", cn, adOpenDynamic, adLockOptimistic, adCmdText
       While Not rs.EOF
          ' AQUI SE VA A CAMBIAR LA RELACION, YA ES DIRECTA
          RSTEMP.Open "SELECT * FROM tfproduc,cargos WHERE consec = caprod AND consec = '" & Trim(rs!Dt_producto) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
          If RSTEMP.BOF And RSTEMP.EOF Then
             MsgBox "EL PRODUCTO CON CLAVE : " & rs!Dt_producto & " DEL TRASLADO CON CLAVE " & rs!DT_CLAVE & " NO EXISTE EN EL CATALOGO DE PRODUCTOS CAJAS: " & rs!dt_cantidad & " PIEZAS: " & rs!dt_cantidad & " PRECIO: " & rs!DT_costo, vbCritical
          Else
              iva = RSTEMP!iva
              ieps = RSTEMP!ieps
              PAQUETES = RSTEMP!PAQUETES
              RSTEMP.Close
              'correccion para poner a precio de venta
              'Importe = rs!DT_cantidad * rs!DT_costo + (rs!DT_cantidadp * (rs!DT_costo / PAQUETES))
              'EN TEORIA ESTA BIEN PERO SE DEBE TOMAR EL PRECIO PZA QUE TRAE LA BD
              'importe = rs!DT_cantidad * rs!DT_venta + (rs!DT_cantidadp * (rs!DT_venta / PAQUETES))
              importe = rs!dt_cantidad * rs!DT_venta + (rs!dt_cantidadp * rs!DT_ventap)
              'Busco dentro de la venta para saber si ya existe el producto
              'tasaieps = validatasa(ieps, iva)
              tasaieps = RSTEMP!tasaieps
              RSTEMP.Open "SELECT * FROM ventas_det WHERE noventa = " & folVta & " AND cl_producto = '" & rs!Dt_producto & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
              If RSTEMP.BOF And RSTEMP.EOF Then
                Stb1.SimpleText = Space(20) & "Agregando la clave del producto " & rs!Dt_producto & " del traslado con clave " & rs!DT_CLAVE
                 'MsgBox "INSERT INTO ventas_det(noventa,cl_producto,cantidad,cantidadp,precosto,precostop,importe,ieps,iva) VALUES (" & folvta & ",'" & rs!dt_producto & "'," & rs!dt_cantidad & "," & rs!dt_cantidadp & "," & rs!dt_costo & "," & rs!dt_costo / paquetes & "," & importe & "," & IEPS & "," & IVA & ")"
                 'cn.Execute "INSERT INTO ventas_det(noventa,cl_producto,cantidad,cantidadp,precio,preciop,importe,ieps,iva,tasaieps) VALUES (" & folVta & ",'" & rs!Dt_producto & "'," & rs!dt_cantidad & "," & rs!dt_cantidadp & "," & rs!DT_costo & "," & rs!DT_costo / Paquetes & "," & importe & "," & IEPS & "," & IVA & "," & tasaieps & ")"
                 cn.Execute "INSERT INTO ventas_det(noventa,cl_producto,cantidad,cantidadp,precio,preciop,importe,ieps,iva,tasaieps,precosto,precostop) VALUES (" & folVta & ",'" & rs!Dt_producto & "'," & rs!dt_cantidad & "," & rs!dt_cantidadp & "," & rs!DT_venta & "," & rs!DT_ventap & "," & importe & "," & ieps & "," & iva & "," & tasaieps & "," & rs!DT_costo & "," & rs!DT_costop & ")"
              Else
                 Stb1.SimpleText = Space(20) & "Actualizando la clave del producto " & rs!Dt_producto & " del traslado con clave " & rs!DT_CLAVE
                 cn.Execute "UPDATE VENTAS_det SET cantidad = cantidad + " & rs!dt_cantidad & ", cantidadp = cantidadp + " & rs!dt_cantidadp & ", Importe = importe + " & importe & " WHERE noventa = " & folVta & " AND cl_producto = '" & rs!Dt_producto & "'"
              End If
          End If
          RSTEMP.Close
          rs.MoveNext
       Wend
       fecha = CStr(Day(AdoTraslados.Recordset!T_FECHA)) & CStr(Month(AdoTraslados.Recordset!T_FECHA)) & CStr(Year(AdoTraslados.Recordset!T_FECHA)) & Trim(AdoTraslados.Recordset!t_sucursalreceptor)
       rs.Close
      End If ' fin de la fecha
    End If ' fin de que sea enviado
    AdoTraslados.Recordset.MoveNext
    Wend
    RSTEMP.Open "SELECT SUM(IMPORTE) AS importe FROM ventas_det WHERE NOVENTA =  " & folVta, cn, adOpenDynamic, adLockOptimistic, adCmdText
    cn.Execute "UPDATE VENTAS SET montototal = " & IIf(IsNull(RSTEMP!importe), 0, RSTEMP!importe) & " WHERE noventa = " & folVta
    RSTEMP.Close
    Stb1.SimpleText = MenAnt
    Stb1.Refresh
    cn.CommitTrans
MsgBox "Proceso Finalizado...", vbInformation, "GENERACION DE VENTA"
End Sub

Function validatasa(TIEPS As Integer, TIVA As Integer) As Integer
validatasa = 0
If TIEPS = 0 And TIVA = 0 Then
   validatasa = 1
ElseIf TIEPS = 0 And TIVA = 15 Then
   validatasa = 2
ElseIf TIEPS = 25 And TIVA = 15 Then
   validatasa = 3
ElseIf TIEPS = 30 And TIVA = 15 Then
   validatasa = 4
ElseIf TIEPS = 50 And TIVA = 15 Then
   validatasa = 5
ElseIf TIEPS = 60 And TIVA = 15 Then
   validatasa = 6
ElseIf TIEPS = 5 And TIVA = 15 Then
   validatasa = 7
Else
   MsgBox "ESTE PRODUCTO NO SE ENCUENTRA DENTRO DE ALGUN DEPARTAMENTO CON TASA REGISTRADA, FAVOR DE INFORMAR AL ADMINISTRADOR DEL SISTEMA", vbCritical, "No se encuentra tasa registrada"
End If
End Function

Private Sub cerrara_Click()
On Error GoTo Error:
AdoDbf.Recordset.Close
MsgBox "Proceso Terminado...", vbInformation
Exit Sub
Error:
   MsgBox "El archivo Se encuentra Cerrado", vbInformation
End Sub

Private Sub cmbFiltro_DblClick()
  'SendKeys "{TAB}"
  keybd_event &H9, 0, 0, 0
End Sub

Private Sub cmbFiltro_LostFocus()
On Error GoTo Error:
 Select Case cmbFiltro.ListIndex
 Case 0  'todos exepto los cancelados
     cCond = "t_motivoCancela Is Null "
     ccondrpt = "{TRASLADOS.t_clave} <> ''"
 Case 1  'Pendientes de recibir
     cCond = "T_enviado = 1 AND T_recibido = 0"
     ccondrpt = "{TRASLADOS.T_enviado} = 1 AND {TRASLADOS.T_recibido} = 0"
 Case 2  'Pendientes recibidos
     cCond = "T_recibido = 1"
     ccondrpt = "{TRASLADOS.T_recibido} = 0"
 Case 3 'TRASLADOS DE FRUTAS
     cCond = "T_FRUTAS = 1"
     ccondrpt = "{TRASLADOS.T_frutas} = 1"
 Case 4 'TRASLADOS DE PANADERIA
     cCond = "T_PAN = 1"
     ccondrpt = "{TRASLADOS.T_pan} = 1"
 Case 5 'TRASLADOS DE MERMAS
     cCond = "T_MERMA = 1"
     ccondrpt = "{TRASLADOS.T_merma} = 1"
 Case 6 'TRASLADOS DE PANADERIA
     cCond = "T_AUTO = 1"
     ccondrpt = "{TRASLADOS.T_auto} = 1"
 End Select
 cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
 
   
 If Sql Then
     cFecha = " AND (month(T_fecha) >= " & Month(dtpFecha(0).Value) & " and (day(T_fecha) > = " & Day(dtpFecha(0).Value) & cOper & " (day(T_fecha)<= " & Day(dtpFecha(1).Value) & " and month(T_fecha)<= " & Month(dtpFecha(1).Value) & ")) and year(T_fecha)>= " & Year(dtpFecha(0).Value) & " and year(T_fecha)<= " & Year(dtpFecha(1).Value) & ")"
  Else
    'COMPATIBILIDAD CON ACCESS
     Dim INI As Date
     Dim FIN As Date
     INI = dtpFecha(0).Value
     FIN = dtpFecha(1).Value
     'PRIMERO MES , DIA , AÑO
    INI = INI - 1
    FIN = FIN + 1
    'cFecha = " AND (month(T_fecha) >= " & Month(dtpfecha(0).value) & " and (day(T_fecha) > = " & Day(dtpfecha(0).value) & cOper & " (day(T_fecha)<= " & Day(dtpfecha(1).value) & " and month(T_fecha)<= " & Month(dtpfecha(1).value) & ")) and year(T_fecha)>= " & Year(dtpfecha(0).value) & " and year(T_fecha)<= " & Year(dtpfecha(1).value) & ")"
    cFecha = " AND  (((TRASLADOS.T_fecha) > #" & Format(INI, "mm /dd/yy") & "#)) AND  (((traslados.t_fecha)<#" & Format(FIN, "mm/dd/yy") & "#))"
     'FIN DE COMPATIBILIDAD
  End If
  
 AdoTraslados.RecordSource = "SELECT * FROM [TRASLADOS] WHERE " & cCond & cFecha
 AdoTraslados.Refresh
 'If Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF) Then AdoTraslados.Recordset.MoveLast
 For N = 0 To 5
   Cmdmoverse(N).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 Next
 'cmdOpcion(0).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(1).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(6).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(4).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 lblInfo.Caption = Str(AdoTraslados.Recordset.RecordCount)
 Exit Sub
Error:
 MsgBox Err.Description
End Sub

Private Sub cmdConAceptar_Click()
If txtContra.Text <> "MODI3456" Then
   MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbExclamation
   txtContra.SetFocus
   SendKeys "+{HOME}"
ElseIf nCancela = 1 Then  'Modificar traslado con opcion de poner en cero el inventario
   Me.fraCon.Visible = False
   'txtCancela.SetFocus
ElseIf nCancela = 0 Then  'Cancelar traslado y regresar la mercancia al inventario
   nOp = 3
   Me.fraCon.Visible = False
   frmTrasladaEnv.Caption = "Modificar envio"
   frmTrasladaEnv.Show
   SendKeys cfolCan
   'SendKeys "{TAB}"
   keybd_event &H9, 0, 0, 0
ElseIf nCancela = 2 Then
   If MsgBox("REALMENTE DESEAS ACTUALIZAR PRECIOS DEL TRASLADO " & AdoTraslados.Recordset!t_clave, vbQuestion + vbYesNo) = vbYes Then
        'If Val(AdoTraslados.Recordset!T_SUCURSALRECEPTOR) = 5 Or Val(AdoTraslados.Recordset!T_SUCURSALRECEPTOR) = 12 Or Val(AdoTraslados.Recordset!T_SUCURSALRECEPTOR) = 13 Or Val(AdoTraslados.Recordset!T_SUCURSALRECEPTOR) = 15 Or Val(AdoTraslados.Recordset!T_SUCURSALRECEPTOR) = 14 Or Val(AdoTraslados.Recordset!T_SUCURSALRECEPTOR) = 27 Then
        If rstSuc!franquicia Then
            cn.Execute "UPDATE DETALLETRASLADO SET DETALLETRASLADO.dt_costo = TFPRODUC.PRECOSTO, DETALLETRASLADO.dt_costoP = TFPRODUC.PRECOSTO /TFPRODUC.paquetes, DETALLETRASLADO.dt_venta = PREPROD.Precio4, DETALLETRASLADO.dt_ventap = PREPROD.Precio4 / TFPRODUC.paquetes FROM PREPROD,TFPRODUC WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & Trim(AdoTraslados.Recordset!t_clave) & "' AND tfproduc.ACTIVO = 1"
        Else 'Precio a tiendas (PRECOSTO de TFPRODUC)
            cn.Execute "UPDATE DETALLETRASLADO SET dt_costo = TFPRODUC.Precosto, dt_costoP = TFPRODUC.Precosto / TFPRODUC.paquetes , dt_venta = PREPROD.precio2, dt_ventaP = PREPROD.precio1 FROM TFPRODUC,PREPROD WHERE PREPROD.preclave = DETALLETRASLADO.dt_producto AND PREPROD.preclave = TFPRODUC.Consec AND TFPRODUC.Consec = DETALLETRASLADO.dt_producto AND DetalleTraslado.dt_clave = '" & Trim(AdoTraslados.Recordset!t_clave) & "' and tfproduc.ACTIVO = 1"
        End If
   MsgBox "LOS PRECIOS SE ACTUALIZARON CORRECTAMENTE", vbInformation
   End If
   fraCon.Visible = False
End If
End Sub


Private Sub cmdConCance_Click()
  fraCon.Visible = False
  ActDesOpc True
End Sub

Private Sub cmdexport_Click()
On Error GoTo Error:
Dim cnFoxPro As ADODB.Connection
Dim sucu As String
Dim cArch As String
frmAreaRecibo.Cmdlg.DialogTitle = "Seleccionar archivo para exportar Traslados"
frmAreaRecibo.Cmdlg.Filter = "Archivos Visual Fox pro (*.dbf) | *.dbf"
frmAreaRecibo.Cmdlg.CancelError = True
frmAreaRecibo.Cmdlg.ShowSave
cRutArc = frmAreaRecibo.Cmdlg.FileName
If cRutArc = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UNA RUTA Y NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
End If
For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
Next
cruta = Mid(cRutArc, 1, nPos)
cArch = Mid(cRutArc, nPos + 1)
cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFile("P:\ENVIOS\ESTTRAS.DBF")
'Set f = fs.GetFile("c:\pitic\ESTTRAS.DBF")
'EN ESTA CONDICION SE DETERMINA SI SE SOBREESCRIBE EL ARCHIVO O BIEN SE ANEXA UN TRASLADO
RESP = MsgBox("Desea Anexar al archivo  " & cArch & "  El traslado Seleccionado ? ", vbYesNo + vbQuestion)
If RESP = vbNo Then
   f.Copy cRutArc, True
End If
sucu = Pedsuc(cArch)
For i = 1 To 30
    TRASLADOSENV(i) = ""
Next
i = 1
AdoDbf.CommandType = adCmdText
AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
AdoDbf.RecordSource = "SELECT * FROM " & cArch
AdoDbf.Refresh
sucut = sucu
MsgBox "Ahora Puede Agregar al Archivo " & arch & " Los Traslados Correspondientes...", vbInformation
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub exporta2()
'On Error GoTo Error:
'TRASLADOSENV(I) = tempotraslado
Set rsttemp = New ADODB.Recordset
' AQUI SOLO VA A GRABAR LOS TRASLADOS QUE HAYA ESCRITO EL OPERADOR
'OTRA CONDICION SERIA POR DIA Y TODOS LOS QUE ESTEN ENVIADOS
tempotraslado = dbgrdTras.Columns(0).Text
rsttemp.Open "SELECT * FROM TRASLADOS WHERE T_CLAVE  = '" & Trim(tempotraslado) & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
If Not rsttemp.EOF Then
   'SE VALIDA QUE SE ENVIE A LA TIENDA CORRECTA
   If sucut <> Trim(rsttemp!t_sucursalreceptor) Then
      MsgBox "Este Traslado No Corresponde a la tienda que selecciono en el Archivo  ", vbInformation
      rsttemp.Close
      Exit Sub
   End If
   traslado = rsttemp!t_clave
   fecha = rsttemp!T_FECHA
   tipo = rsttemp!t_tipo
   perenv = rsttemp!t_perenv
   perfle = rsttemp!t_perfle
   COSTOT = rsttemp!t_costo
   emisor = rsttemp!t_sucursalemisor
   Pedido = rsttemp!t_pedido
   receptor = rsttemp!t_sucursalreceptor
   Folio = rsttemp!t_foliotie
Else
   MsgBox "No Existe Traslado " & tempotraslado, vbInformation
End If
rsttemp.Close
rsttemp.Open "SELECT * FROM DETALLETRASLADO WHERE DT_CLAVE  = '" & Trim(tempotraslado) & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
While Not rsttemp.EOF
      AdoDbf.Recordset.AddNew
      AdoDbf.Recordset!traslado = Trim(traslado)
      AdoDbf.Recordset!fecha = fecha
      If tipo Then
         AdoDbf.Recordset!tipo = "1"
      Else
         AdoDbf.Recordset!tipo = "0"
      End If
      AdoDbf.Recordset!perenv = perenv
      AdoDbf.Recordset!perfle = perfle
      AdoDbf.Recordset!COSTOT = IIf(IsNull(COSTOT), 0, COSTOT)
      AdoDbf.Recordset!emisor = emisor
      AdoDbf.Recordset!receptor = receptor
      AdoDbf.Recordset!Folio = IIf(IsNull(Folio), 0, Folio)
      AdoDbf.Recordset!producto = IIf(IsNull(rsttemp!Dt_producto), "", Trim(rsttemp!Dt_producto))
      AdoDbf.Recordset!cantidad = rsttemp!dt_cantidad
      AdoDbf.Recordset!cantidadp = rsttemp!dt_cantidadp
      AdoDbf.Recordset!Pedido = IIf(IsNull(rsttemp!dt_pedido), "", Trim(rsttemp!dt_pedido))
      AdoDbf.Recordset!costo = IIf(IsNull(rsttemp!DT_costo), 0, Trim(rsttemp!DT_costo))
      AdoDbf.Recordset!costop = IIf(IsNull(rsttemp!DT_costop), 0, Trim(rsttemp!DT_costop))
      AdoDbf.Recordset!importe = IIf(IsNull(rsttemp!DT_IMPORTE), 0, Trim(rsttemp!DT_IMPORTE))
      AdoDbf.Recordset!venta = IIf(IsNull(rsttemp!DT_venta), 0, Trim(rsttemp!DT_venta))
      AdoDbf.Recordset!ventaP = IIf(IsNull(rsttemp!DT_ventap), 0, Trim(rsttemp!DT_ventap))
      AdoDbf.Recordset!Importado = False
      AdoDbf.Recordset.Update
      rsttemp.MoveNext
      nreg = nreg + 1
   Wend
rsttemp.Close
MsgBox "El traslado se agrego correctamente", vbInformation
Exit Sub
Error:
  MsgBox "El archivo no ha sido abierto ", vbCritical, "ERROR"
End Sub

Private Sub actualizasugerido(traslado As String)
'TRASLADO = AdoDbf.Recordset!TRASLADO
Set rstbusca1 = New ADODB.Recordset
fechatra = AdoDbf.Recordset!fecha
tipo = AdoDbf.Recordset!tipo
perenv = AdoDbf.Recordset!perenv
perfle = AdoDbf.Recordset!perfle
COSTOT = AdoDbf.Recordset!COSTOT
Pedido = AdoDbf.Recordset!Pedido
emisor = AdoDbf.Recordset!emisor
receptor = AdoDbf.Recordset!receptor
Folio = AdoDbf.Recordset!Folio
While (AdoDbf.Recordset.EOF = False) And ("m" & Trim(AdoDbf.Recordset!traslado) = Trim(traslado))
          
          producto = AdoDbf.Recordset!producto
          cantidad = AdoDbf.Recordset!cantidad
          cantidadp = AdoDbf.Recordset!cantidadp
          Pedido = AdoDbf.Recordset!Pedido
          costo = AdoDbf.Recordset!costo
          costop = AdoDbf.Recordset!costop
          importe = AdoDbf.Recordset!importe
          venta = AdoDbf.Recordset!venta
          ventaP = AdoDbf.Recordset!ventaP
          'SE VERIFICA QUE EXISTA EL PRODUCTO EN CASO CONTRARIO SE ANEXA
          CAD = " SELECT * FROM DETALLEFACTURA  WHERE DF_PEDIDO  = '" & Trim(Pedido) & "' and df_prod = '" & Trim(producto) & "'"
          rstbusca1.Open CAD, cn, adOpendinamic, adLockOptimistic, adCmdText
          If rstbusca1.EOF Then
             CAD = "INSERT INTO DETALLEFACTURA(DF_PROD,DF_PEDIDO,DF_CANTSOL,DF_CANTSOLP,DF_COSTO,DF_CANTREAL,DF_CANTREALP) VALUES( " & _
                   "'" & Trim(producto) & "','" & Trim(Pedido) & "'," & cantidad & "," & cantidadp & "," & costo & "," & cantidad & "," & cantidadp & ")"
          Else
             CAD = "update detallefactura  set df_cantreal = " & cantidad & ", df_costo = " & costo & ",df_cantrealp = " & cantidadp & " where df_pedido = '" & Trim(Pedido) & "' and df_prod = '" & Trim(producto) & "'"
          End If
          rstbusca1.Close
          'MsgBox cad
          cn.Execute CAD
          COSTOT = costo + costop
          AdoDbf.Recordset.MoveNext
          'EN EL CASO DE QUE HAYA LLEGADO AL FINAL DEL ARCHIVO
          If AdoDbf.Recordset.EOF = True Then
             MsgBox "Proceso de Importacion Finalizado...", vbInformation
             Exit Sub
          End If
Wend
End Sub

Private Sub grabatraslado()
'SE DEBE VALIDAR QUE NO EXISTA
Set rstbusca = New ADODB.Recordset
fechatra = AdoDbf.Recordset!fecha
tipo = AdoDbf.Recordset!tipo
traslado = AdoDbf.Recordset!traslado
perenv = AdoDbf.Recordset!perenv
perfle = AdoDbf.Recordset!perfle
COSTOT = AdoDbf.Recordset!COSTOT
Pedido = AdoDbf.Recordset!Pedido
emisor = AdoDbf.Recordset!emisor
receptor = AdoDbf.Recordset!receptor
Folio = AdoDbf.Recordset!Folio
rstbusca.Open " Select * from traslados where t_clave = '" & Trim(traslado) & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
If rstbusca.EOF Then
        CAD = "INSERT INTO traslados(t_recibido,t_clave,t_fecha,t_tipo,t_perenv,t_perfle,t_costo,t_pedido,t_sucursalemisor,t_sucursalreceptor,t_enviado,t_foliotie,t_entrada) " & _
        "values ( 1, '" & Trim(traslado) & "','" & fechatra & "'," & 1 & "," & perenv & "," & perfle & "," & COSTOT & ",'" & Pedido & "'," & emisor & "," & receptor & "," & 0 & "," & Folio & ",1)"
        'MsgBox cad
        cn.Execute CAD
Else
        soloactualizar = True
End If
rstbusca.Close
'SE GENERA EL DETALLE DEL TRASLADO
While (AdoDbf.Recordset.EOF = False) And (Trim(AdoDbf.Recordset!traslado) = Trim(traslado))
          producto = AdoDbf.Recordset!producto
          cantidad = AdoDbf.Recordset!cantidad
          cantidadp = AdoDbf.Recordset!cantidadp
          Pedido = AdoDbf.Recordset!Pedido
          costo = AdoDbf.Recordset!costo
          costop = AdoDbf.Recordset!costop
          importe = AdoDbf.Recordset!importe
          venta = AdoDbf.Recordset!venta
          ventaP = AdoDbf.Recordset!ventaP
          CAD = "INSERT INTO DETALLETRASLADO(DT_CLAVE,DT_PRODUCTO,DT_CANTIDAD,DT_CANTIDADP,DT_PEDIDO,DT_COSTO,DT_COSTOP,DT_IMPORTE,DT_VENTA,DT_VENTAP)" & _
                "  VALUES ( '" & Trim(traslado) & "','" & Trim(producto) & "'," & cantidad & "," & cantidadp & ",'" & Trim(Pedido) & "'," & costo & "," & costop & "," & importe & "," & venta & "," & ventaP & ")"
          'MsgBox cad
          If soloactualizar = False Then
             cn.Execute CAD
          End If
          COSTOT = costo + costop
          AdoDbf.Recordset.MoveNext
          'EN EL CASO DE QUE HAYA LLEGADO AL FINAL DEL ARCHIVO
          If AdoDbf.Recordset.EOF = True Then
             MsgBox "Proceso de Importacion Finalizado...", vbInformation
             Exit Sub
          End If
Wend
End Sub

Private Sub cmdimporta_Click()
On Error GoTo Error:
Dim cArch  As String
MenAnt = Stb1.SimpleText
Set rstbusca = New ADODB.Recordset
frmAreaRecibo.Cmdlg.FileName = ""
frmAreaRecibo.Cmdlg.Filter = "Archivos Dbase (*.dbf) | *.dbf"
frmAreaRecibo.Cmdlg.ShowOpen
cRutArc = frmAreaRecibo.Cmdlg.FileName
If cRutArc = "" Or IsNull(cRutArc) Then
    MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
    Exit Sub
End If
For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
Next
cruta = Mid(cRutArc, 1, nPos)
cArch = Mid(cRutArc, nPos + 1)
cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.
AdoDbf.CommandType = adCmdText
AdoDbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties='DSN=PITICODBF;UID=;SourceDB=" & cruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=SPANISH;Null=Sí;Deleted=Sí;';Initial Catalog= " & cruta
AdoDbf.RecordSource = "SELECT * FROM " & cArch & " order BY traslado"
AdoDbf.Refresh
If AdoDbf.Recordset.BOF And AdoDbf.Recordset.EOF Then
   MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
   Exit Sub
End If
If AdoDbf.Recordset!Importado Then
      MsgBox "EL ARCHIVO SELECCIONADO YA FUE IMPORTADO", vbInformation
      Exit Sub
End If
   
traslado = "moy"
SUC = Pedsuc(cArch)
If IsNull(SUC) Or Trim(SUC) = "" Then
      MsgBox "EL NOMBRE DEL ARCHIVO: " & cArch & "NO ESTA REGISTRADO EN EL SISTEMA" & Chr(13) _
      & "Y NO SE LE ASIGNARA SUCURSAL AL PEDIDO" & Chr(13) & "FAVOR DE AVISAR AL ADMINISTRADOR DEL SISTEMA", vbCritical
      Exit Sub
End If
Dim Pedido As String
While Not AdoDbf.Recordset.EOF
   COSTOT = 0
   If Not Trim(AdoDbf.Recordset!traslado) = Trim(traslado) Then
        'SE GENERAN LOS DATOS GENERALES
        traslado = AdoDbf.Recordset!traslado
        fechatra = AdoDbf.Recordset!fecha
        tipo = AdoDbf.Recordset!tipo
        perenv = AdoDbf.Recordset!perenv
        perfle = AdoDbf.Recordset!perfle
        COSTOT = AdoDbf.Recordset!COSTOT
        Pedido = AdoDbf.Recordset!Pedido
        emisor = AdoDbf.Recordset!emisor
        receptor = AdoDbf.Recordset!receptor
        Folio = AdoDbf.Recordset!Folio
        'condiciones para determinar si se graba en pedidos o en traslados
'        If Not IsNull(Pedido) Then
'           If checasugerido(Pedido) Then
'              ' se acualiza la fecha de recibo
'              cad = "update pedidos set p_fecentreal = '" & fecha & "', p_recibido = 1 where p_pedido = '" & Trim(Pedido) & "'"
'              cn.Execute cad
'              actualizasugerido (TRASLADO)
'           Else
'              'SE TRATA DE UN SUGERIDO PERO NO EXISTE EN LA BASE DE RECIBO
'              grabatraslado
'           End If
'        Else
           grabatraslado
'        End If
   End If
   Stb1.SimpleText = MenAnt
   Stb1.Refresh
   'traslado = AdoDbf.Recordset!traslado
   If AdoDbf.Recordset.EOF = True Then
             MsgBox "Proceso de Importacion Finalizado...", vbInformation
             Exit Sub
   End If
Wend
Stb1.Refresh
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Function checasugerido(numped As String) As Boolean
Set rstbusca = New ADODB.Recordset
rstbusca.Open " Select * from pedidos where p_pedido  = '" & Trim(numped) & "'", cn, adOpenStatic, adLockOptimistic, adCmdText
If rstbusca.EOF Then
   checasugerido = False
Else
   checasugerido = True
End If
End Function

Private Sub cmdMoverse_Click(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0  'Primer registro
    AdoTraslados.Recordset.MoveFirst
Case 1  ' Anterior
    AdoTraslados.Recordset.MovePrevious
    If AdoTraslados.Recordset.BOF Then AdoTraslados.Recordset.MoveFirst
Case 2  ' Siguiente
    AdoTraslados.Recordset.MoveNext
    If AdoTraslados.Recordset.EOF Then AdoTraslados.Recordset.MoveLast
Case 3  'Ultimo
    AdoTraslados.Recordset.MoveLast
Case 4  'Buscar clave de ltraslado
    cCve = InputBox("Introduzca la clave del traslado a buscar", "Introducir clave")
    If IsNull(Trim(cCve)) Or Trim(cCve) = "" Then Exit Sub
    Antes = dbgrdTras.Bookmark
    AdoTraslados.Recordset.MoveFirst
    AdoTraslados.Recordset.Find "T_clave = '" & Trim(cCve) & "'"
    If AdoTraslados.Recordset.EOF Then
        MsgBox "LA CLAVE " & UCase(cCve) & " NO SE ENCUENTRA EN LOS TRASLADOS " + IIf(Me.cmbFiltro.Text = "TODOS", "" & Chr(13) & " EN EL PERIODO SELECCIONADO", cmbFiltro.Text), vbExclamation
        dbgrdTras.Bookmark = Antes
    End If
Case 5
    cMensaje = Stb1.SimpleText
    Stb1.SimpleText = Space(90) + "Espere un momento generando reporte..."
    Stb1.Refresh
    cFecharpt = " AND {TRASLADOS.t_fecha} >= Date(" & CStr(Year(dtpFecha(0).Value)) & "," & CStr(Month(dtpFecha(0).Value)) & "," & CStr(Day(dtpFecha(0).Value)) & ") AND {TRASLADOS.t_fecha} <= Date(" & CStr(Year(dtpFecha(1).Value)) & "," & CStr(Month(dtpFecha(1).Value)) & "," & CStr(Day(dtpFecha(1).Value)) & ")"
    cr1.Connect = cCadConex
    cr1.ReportFileName = App.Path & "\Traslada.rpt"
    cr1.WindowTitle = "Reporte de traslados"
    cr1.Formulas(0) = "FORMSELEC = " & ccondrpt & cFecharpt
    cr1.Formulas(1) = "TRASLADO = 'LISTADO DE TRASLADOS " & IIf(cmbFiltro.Text = "TODOS", "", cmbFiltro.Text) & " DEL " & Trim(dtpFecha(0).Value) & " AL " & Trim(dtpFecha(1).Value) & " '"
    cr1.Action = 1
    Stb1.SimpleText = cMensaje
    Stb1.Refresh
Case 6
     cfolCan = dbgrdTras.Columns(0).Text
     nCancela = 2
     fraCon.Visible = True
     txtContra.Text = ""
     txtContra.SetFocus
     AdoTraslados.Recordset.Find "T_CLAVE = '" & cfolCan & "'"
End Select
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdopcion_Click(Index As Integer)
On Error GoTo Error:
Select Case Index
Case 0 'Nuevo traslado
     cModo = ""   'ENVIO a tiendas
     nOp = 1    'Nuevo envio
     Forma = 0  'Salida de productos   (Disminuye el inventario)
     frmTrasladaEnv.Caption = "Capturar nuevo envio"
     frmTrasladaEnv.Show
Case 1 'Modificar traslado
     If AdoTraslados.Recordset!t_entrada Then
        MsgBox "NO PUEDE MODIFICAR EL TRASLADO EN SALIDA PORQUE ES UNA ENTRADA" & Chr(13) & "SELECCIONE LA OPCION MOD. REC.", vbInformation
        Exit Sub
     End If
     cModo = ""   'ENVIO a tiendas
     nOp = 0     'Modificaciones
     Forma = 0   'Salida de productos
     strcveprod = Me.dbgrdTras.Columns(0).Text
     frmTrasladaEnv.Caption = "Modificar envio"
     
     frmTrasladaEnv.Show
     SendKeys dbgrdTras.Columns(0).Text
     keybd_event &H9, 0, 0, 0
     
Case 2 'Nueva Recepcion del traslado
     cModo = ""   'ENVIO a tiendas
     nOp = 1    ' Es un nuevo traslado
     Forma = 1  'Es una recepcion de productos  (Incrementa el inventario)
     frmTrasladaEnv.Caption = "Recibir envio"
     frmTrasladaEnv.Show
Case 6
     If Not AdoTraslados.Recordset!t_entrada Then
        MsgBox "NO PUEDE MODIFICAR EL TRASLADO EN ENTRADA PORQUE ES UNA SALIDA" & Chr(13) & "SELECCIONE LA OPCION MOD. ENV.", vbInformation
        Exit Sub
     End If
     nOp = 0
     Forma = 1
     frmTrasladaEnv.Caption = "Modificar recepcion de traslado"
     frmTrasladaEnv.Show
     SendKeys Me.dbgrdTras.Columns(0).Text
     keybd_event &H9, 0, 0, 0
Case 3
     Unload Me
     If nOp = 30 Then
           frmAreaRecibo.Show
     End If
Case 4 'Cancelar traslado
     If AdoTraslados.Recordset!t_entrada Then
        MsgBox "NO ES POSIBLE CANCELAR EL TRASLADO PORQUE ES UNA ENTRADA", vbInformation
        Exit Sub
     End If
     nCancela = 0
     cfolCan = dbgrdTras.Columns(0).Text
     ActDesOpc False 'Procedimiento que desactiva la opcion de menus
     fraCon.Visible = True
     txtContra.Text = ""
     txtContra.SetFocus
Case 5  'Modificar
    nCancela = 1
    cfolCan = dbgrdTras.Columns(0).Text
    ActDesOpc False 'Procedimiento que desactiva la opcion de menus
    fraCon.Visible = True
    txtContra.Text = ""
    txtContra.SetFocus
Case 7
    cModo = "DEVO"   'Devolucion a proveedores
    nOp = 1    'Nuevo envio
    Forma = 0  'Salida de productos   (Disminuye el inventario)
    frmTrasladaEnv.Caption = "Devolucion a proveedores"
    frmTrasladaEnv.Show
Case 8
    cModo = "DEVO"   'ENVIO a tiendas
    nOp = 0     'Modificaciones
    Forma = 0   'Salida de productos
    frmTrasladaEnv.Caption = "Modificar Devolucion"
    frmTrasladaEnv.Show
    SendKeys Me.dbgrdTras.Columns(0).Text
    keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
End Select
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub cmdopcion_GotFocus(Index As Integer)
If Index = 0 Then Unload frmAreaRecibo
End Sub

Private Sub cambiatienda()
fratda.Enabled = True
fratda.Visible = True
Set rstSuc = New ADODB.Recordset
  rstSuc.CursorType = adOpenDynamic
  rstSuc.Source = "SELECT * FROM  CatTienda"
  rstSuc.ActiveConnection = cCadConex
  rstSuc.Open
  cmbFiltrotda.Clear
  cmbFiltrotda.AddItem "ATZOMPA  13"
  cmbFiltrotda.AddItem "MIAHUATLAN  15"
  cmbFiltrotda.AddItem "TLACOLULA  14"
  cmbFiltrotda.AddItem "ZIMATLAN  27"
  cmbFiltrotda.AddItem "TODOS"
  cmbFiltrotda.ListIndex = 4
End Sub


Private Sub CmdRefresh_Click()
  AdoTraslados.Refresh
End Sub

Private Sub cmdRegresar_Click()
  FraCancelar.Visible = False
  ActDesOpc True
End Sub


Private Sub dbgrdTras_DblClick()
If Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF) Then cmdopcion_Click 1
End Sub


Private Sub dbgrdTras_HeadClick(ByVal ColIndex As Integer)
  Me.Stb1.SimpleText = Space(65) + "Espere un momento ordenando Pedidos por " & dbgrdTras.Columns(ColIndex).Caption
  AdoTraslados.RecordSource = "SELECT * FROM [Traslados] WHERE " & cCond & cFecha & " ORDER BY " & dbgrdTras.Columns(ColIndex).DataField
  AdoTraslados.Refresh
  Stb1.SimpleText = Space(85) + "Pedidos ordenandos por " & dbgrdTras.Columns(ColIndex).Caption
End Sub

Private Sub dbgrdTras_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
   PopupMenu Me.x
End If

End Sub

Private Sub dbgrdTras_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
 If dbgrdTras.SelBookmarks.Count > 0 Then dbgrdTras.SelBookmarks.Remove 0
 dbgrdTras.SelBookmarks.Add dbgrdTras.RowBookmark(dbgrdTras.Row)
End Sub

Private Sub dtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then keybd_event &H9, 0, 0, 0
End Sub

Private Sub eliminart_Click()
MsgBox "Por el Momento no se pueden Eliminar, vuelva a Realizar el Proceso", vbInformation
End Sub

Private Sub fe_Click()
mfec = InputBox("Escriba la Fecha Correcta del Sistema", "CAMBIO FECHA")
clave = Me.AdoTraslados.Recordset!t_clave
CAD = "update traslados set t_fecha = '" & Trim(mfec) & "' where t_clave = '" & Trim(clave) & "'"
'MsgBox cad
cn.Execute CAD
Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo Error:
Open App.Path & "\envfac.TXT" For Output As #1
  Set rstSuc = New ADODB.Recordset
  rstSuc.CursorType = adOpenDynamic
  rstSuc.Source = "SELECT * FROM  CatTienda"
  rstSuc.ActiveConnection = cCadConex
  rstSuc.Open
  cmbFiltro.AddItem "TODOS"
  cmbFiltro.AddItem "PENDIENTES DE RECIBIR"
  cmbFiltro.AddItem "RECIBIDOS"
  cmbFiltro.AddItem "FRUTAS"
  cmbFiltro.AddItem "PANADERIA"
  cmbFiltro.AddItem "MERMAS"
  cmbFiltro.AddItem "AUTOCONSUMO"
  cmbFiltro.ListIndex = 0
  If dtpFecha(0).Value = "01/01/02" Then dtpFecha(0).Value = date
  If dtpFecha(1).Value = "01/01/02" Then dtpFecha(1).Value = date
 
  cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
  
  If Sql Then
     cFecha = " AND (month(T_fecha) >= " & Month(dtpFecha(0).Value) & " and (day(T_fecha) > = " & Day(dtpFecha(0).Value) & cOper & " (day(T_fecha)<= " & Day(dtpFecha(1).Value) & " and month(T_fecha)<= " & Month(dtpFecha(1).Value) & ")) and year(T_fecha)>= " & Year(dtpFecha(0).Value) & " and year(T_fecha)<= " & Year(dtpFecha(1).Value) & ")"
  Else
    'COMPATIBILIDAD CON ACCESS
     Dim INI As Date
     Dim FIN As Date
     INI = dtpFecha(0).Value
     FIN = dtpFecha(1).Value
     'PRIMERO MES , DIA , AÑO
    INI = INI - 1
    FIN = FIN + 1
    'cFecha = " AND (month(T_fecha) >= " & Month(dtpfecha(0).value) & " and (day(T_fecha) > = " & Day(dtpfecha(0).value) & cOper & " (day(T_fecha)<= " & Day(dtpfecha(1).value) & " and month(T_fecha)<= " & Month(dtpfecha(1).value) & ")) and year(T_fecha)>= " & Year(dtpfecha(0).value) & " and year(T_fecha)<= " & Year(dtpfecha(1).value) & ")"
    cFecha = " AND  (((TRASLADOS.T_fecha) > #" & Format(INI, "mm /dd/yy") & "#)) AND  (((traslados.t_fecha)<#" & Format(FIN, "mm/dd/yy") & "#))"
     'FIN DE COMPATIBILIDAD
  End If
  
  
  cCond = "t_motivoCancela is Null "                  ' Filtro por default todos los pedidos
  ccondrpt = "{TRASLADOS.t_clave} <> '' "    ' Filtro por default del RPT

  AdoTraslados.ConnectionString = cCadConex
  AdoTraslados.CommandType = adCmdText
  AdoTraslados.RecordSource = "SELECT * FROM [Traslados] WHERE " & cCond & cFecha
  AdoTraslados.Refresh
  'If Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF) Then AdoTraslados.Recordset.MoveLast
  For N = 0 To 5
     Cmdmoverse(N).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
  Next
 'cmdOpcion(0).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(1).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(6).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(4).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 lblInfo.Caption = Str(AdoTraslados.Recordset.RecordCount)

  nOp = 30 'Para saber cuando se cargan las formas
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub Form_Activate()
On Error GoTo Error:
 'ActDesOpc True
 nOp = 30   'para saber que forma se carga ,se descarga el frmtraslados solamnete cuando es 30
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Close #1
If nOp = 30 Then
  frmAreaRecibo.Show
End If
End Sub

Private Sub ActDesOpc(lvalor As Boolean)
  For N = 0 To 5
     cmdopcion(N).Enabled = lvalor
  Next
End Sub

Private Sub mnufacturar_Click()
Print #1, Me.AdoTraslados.Recordset!t_clave
MsgBox "Se ha agregado al archivo de Facturas el envio " & AdoTraslados.Recordset!t_clave
End Sub

Private Sub qtafac_Click()
resp1 = MsgBox("Este proceso solo se hara con el traslado Activo, Deseas Continuar ?  " & vbCrLf, vbYesNo)
If resp1 = vbYes Then
    CAD = "update traslados set t_venta = 0 where t_clave = '" & Me.AdoTraslados.Recordset!t_clave & "'"
    cn.Execute CAD
    AdoTraslados.Refresh
End If
End Sub

Private Sub txtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
  Case 27
       FraGen.Visible = False
       FraImp.Visible = False
       ActDesOpc True
  Case 13
       KeyAscii = 0
       'SendKeys vbTab
       keybd_event &H9, 0, 0, 0: keybd_event &H9, 0, &H2, 0
End Select
End Sub

Private Sub txtCancela_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
   If Len(txtCancela.Text) <> 15 Then
      txtCancela.SetFocus
      Exit Sub
   End If
   'Busco en el catalogo de productos el codigo
   Set rsttemp = New ADODB.Recordset
   rsttemp.Open "SELECT * FROM tfproduc WHERE Consec = '" & Trim(Mid(txtCancela.Text, 1, 10)) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rsttemp.BOF And rsttemp.EOF Then
        MsgBox "EL ARTICULO CON LA CLAVE:" & Trim(Mid(txtCancela.Text, 1, 10)) & " NO EXISTE EN EL CATALOGO DE PRODUCTOS", vbExclamation
        txtCancela.Text = ""
        txtCancela.SetFocus
        Exit Sub
   End If
   nProd = rsttemp!PAQUETES
   'Busco la etiqueta
   lbletiquetas(8).Caption = rsttemp!descripc & Chr(13) & CStr(rsttemp!PAQUETES) & " X " & CStr(rsttemp!PAQUETES) & " " & rsttemp!medida
   rsttemp.Close
   rsttemp.Open "SELECT * FROM CODIGOS WHERE Codigo = '" & txtCancela.Text & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rsttemp.BOF And rsttemp.EOF Then
      MsgBox "LA ETIQUETA CON LA CLAVE:" & txtCancela.Text & " NO EXISTE", vbExclamation
      txtCancela.Text = ""
      txtCancela.SetFocus
      Exit Sub
   End If
   
   If optopc(0).Value = True Then   'Entrada de inventario
      If rsttemp!Entradainv = "0" Then
         MsgBox "LA ETIQUETA ESPECIFICADA AUN NO HA ENTRADO AL INVENTARIO", vbCritical
         txtCancela.Text = ""
         txtCancela.SetFocus
         Exit Sub
      ElseIf rsttemp!Salidainv = "1" Then
         MsgBox "LA ETIQUETA ESPECIFICADA YA SALIO DEL INVENTARIO" & Chr(13) & "CON FECHA Y HORA SIGUIENTE: " & rsttemp!FechaSalida & Chr(13) & "POR LO TANTO NO SE PUEDE CANCELAR LA ENTRADA", vbCritical
         txtCancela.Text = ""
         txtCancela.SetFocus
         Exit Sub
      ElseIf MsgBox("LA ETIQUETA ESPECIFICADA ENTRO AL INVENTARIO" & Chr(13) & "CON FECHA Y HORA SIGUIENTE: " & rsttemp!fechaEntrada & Chr(13) & Chr(13) & "REALMENTE DESEAS CANCELARLA Y DISMINUIR EL INVENTARIO", vbQuestion + vbYesNo) = vbYes Then
         rsttemp!MotivoCancela = txtcampos(6).Text
         rsttemp!FechaCancela = date + Time
         rsttemp!fechaEntrada = Null
         rsttemp!Entradainv = "0"
         rsttemp.Update
         cn.Execute "UPDATE Inventario SET inCant = Incant - " & CStr(nProd) & " WHERE inprod = '" & Mid(txtCancela.Text, 1, 10) & "'"
      End If
   ElseIf optopc(1).Value = True Then   'Salida de inventario
      If rsttemp!Salidainv = "0" Then
         MsgBox "LA ETIQUETA ESPECIFICADA AUN NO HA SALIDO DEL INVENTARIO", vbCritical
         txtCancela.Text = ""
         txtCancela.SetFocus
         Exit Sub
      ElseIf MsgBox("LA ETIQUETA ESPECIFICADA SALIO DEL INVENTARIO EN EL TRASLADO: " & rsttemp!traslado & Chr(13) & "CON FECHA Y HORA SIGUIENTE: " & rsttemp!FechaSalida & Chr(13) & Chr(13) & "REALMENTE DESEAS CANCELARLA Y AUMENTAR EL INVENTARIO", vbQuestion + vbYesNo) = vbYes Then
         rsttemp!MotivoCancela = txtcampos(6).Text
         rsttemp!FechaCancela = date + Time
         rsttemp!FechaSalida = Null
         rsttemp!Salidainv = "0"
         rsttemp.Update
         cn.Execute "UPDATE Inventario SET inCant = Incant + " & CStr(nProd) & " WHERE inprod = '" & Mid(txtCancela.Text, 1, 10) & "'"
         If Not IsNull(rsttemp!traslado) Or Not Trim(rsttemp!traslado) = "" Then
            cn.Execute "UPDATE DetalleTraslado SET dt_cantidad = dt_cantidad - 1 WHERE dt_producto = '" & Mid(txtCancela.Text, 1, 10) & "' AND dt_clave = '" & rsttemp!traslado & "'"
         End If
      End If
   End If
   txtCancela.Text = ""
   
End If
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub txtContra_GotFocus()
AdoTraslados.Recordset.MoveFirst
AdoTraslados.Recordset.Find "t_clave = '" & cfolCan & "'"
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdConAceptar_Click
ElseIf KeyAscii = 27 Then
   cmdConCance_Click
End If
End Sub

Private Sub dtpFecha_LostFocus(Index As Integer)
On Error GoTo Error:
 cOper = IIf(Month(dtpFecha(0).Value) = Month(dtpFecha(1).Value), " AND ", " OR ")
 If Sql Then
     cFecha = " AND (month(T_fecha) >= " & Month(dtpFecha(0).Value) & " and (day(T_fecha) > = " & Day(dtpFecha(0).Value) & cOper & " (day(T_fecha)<= " & Day(dtpFecha(1).Value) & " and month(T_fecha)<= " & Month(dtpFecha(1).Value) & ")) and year(T_fecha)>= " & Year(dtpFecha(0).Value) & " and year(T_fecha)<= " & Year(dtpFecha(1).Value) & ")"
  Else
    'COMPATIBILIDAD CON ACCESS
     Dim INI As Date
     Dim FIN As Date
     INI = dtpFecha(0).Value
     FIN = dtpFecha(1).Value
     'PRIMERO MES , DIA , AÑO
    INI = INI - 1
    FIN = FIN + 1
    'cFecha = " AND (month(T_fecha) >= " & Month(dtpfecha(0).value) & " and (day(T_fecha) > = " & Day(dtpfecha(0).value) & cOper & " (day(T_fecha)<= " & Day(dtpfecha(1).value) & " and month(T_fecha)<= " & Month(dtpfecha(1).value) & ")) and year(T_fecha)>= " & Year(dtpfecha(0).value) & " and year(T_fecha)<= " & Year(dtpfecha(1).value) & ")"
    cFecha = " AND  (((TRASLADOS.T_fecha) > #" & Format(INI, "mm /dd/yy") & "#)) AND  (((traslados.t_fecha)<#" & Format(FIN, "mm/dd/yy") & "#))"
     'FIN DE COMPATIBILIDAD
  End If
  

 AdoTraslados.RecordSource = "SELECT * FROM [Traslados] WHERE " & cCond & cFecha & " ORDER BY T_fecha"
 AdoTraslados.Refresh
 lblInfo.Caption = Str(AdoTraslados.Recordset.RecordCount)
 For N = 0 To 5  'Si esta vacio el recordset desactivo las opciones
   Cmdmoverse(N).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 Next
 'cmdOpcion(0).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(1).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 'cmdOpcion(2).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(4).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
 cmdopcion(6).Enabled = Not (AdoTraslados.Recordset.BOF And AdoTraslados.Recordset.EOF)
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub x_Click()
'If KeyAscii = 113 Then
   
'End If
End Sub
