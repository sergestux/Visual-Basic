VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAreaRecibo 
   Caption         =   "TIENDAS"
   ClientHeight    =   8490
   ClientLeft      =   150
   ClientTop       =   -75
   ClientWidth     =   12120
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000D&
   Icon            =   "frmAreaRecibo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   12120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraUtiOfi 
      Height          =   4695
      Left            =   1560
      TabIndex        =   53
      Top             =   1800
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame FRArutas 
         Height          =   2895
         Left            =   3240
         TabIndex        =   86
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CommandButton cmdGenerar 
            Caption         =   "&Generar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   88
            Top             =   2400
            Width           =   2295
         End
         Begin VB.ListBox lstrutas 
            BackColor       =   &H00C0FFFF&
            Height          =   2085
            ItemData        =   "frmAreaRecibo.frx":030A
            Left            =   120
            List            =   "frmAreaRecibo.frx":0329
            Style           =   1  'Checkbox
            TabIndex        =   87
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.CommandButton Cmdutiofi 
         Caption         =   "Inf. Preventistas"
         Height          =   525
         Index           =   4
         Left            =   7320
         Picture         =   "frmAreaRecibo.frx":0432
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Genera información para preventistas"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Cmdlistagte 
         Caption         =   "&Listado Agente"
         Height          =   400
         Left            =   6360
         TabIndex        =   82
         ToolTipText     =   "Listado de precios para agentes"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Frame fraperi 
         Caption         =   "Período de Generación..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1680
         TabIndex        =   65
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
         Begin VB.CommandButton Cmdopt 
            Caption         =   "&Procesar"
            Height          =   375
            Index           =   19
            Left            =   1680
            TabIndex        =   67
            Top             =   1080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSMask.MaskEdBox txtinicio 
            Height          =   375
            Left            =   480
            TabIndex        =   66
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtfinal 
            Height          =   375
            Left            =   3120
            TabIndex        =   68
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin VB.Label linicio 
            Alignment       =   2  'Center
            Caption         =   "Fecha Inicial"
            Height          =   255
            Left            =   480
            TabIndex        =   70
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lfinal 
            Alignment       =   2  'Center
            Caption         =   "Fecha Final"
            Height          =   255
            Left            =   3120
            TabIndex        =   69
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CommandButton Cmdutiofi 
         Caption         =   "&Cambios"
         Height          =   525
         Index           =   1
         Left            =   2040
         Picture         =   "frmAreaRecibo.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Generar archivo con cambio de precios"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Cmdutiofi 
         Caption         =   "&Actualizar"
         Height          =   525
         Index           =   0
         Left            =   240
         Picture         =   "frmAreaRecibo.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Actualizar base de datos desde Dbf"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdutiTer 
         Caption         =   "&Terminar"
         Height          =   400
         Left            =   7560
         TabIndex        =   62
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdverofertas 
         Caption         =   "List. de &Ofertas "
         Height          =   400
         Left            =   360
         TabIndex        =   61
         ToolTipText     =   "Listado de Ofertas"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdliscam 
         Caption         =   "List. cambios"
         Height          =   400
         Left            =   1560
         TabIndex        =   60
         ToolTipText     =   "Listado de cambios"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton Cmdutiofi 
         Caption         =   "Catálogo Comple&to"
         Height          =   525
         Index           =   3
         Left            =   5640
         Picture         =   "frmAreaRecibo.frx":0738
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Genera archivo con todo el cátalogo de productos"
         Top             =   240
         Width           =   1455
      End
      Begin ComctlLib.ProgressBar probar1o 
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   4200
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton Cmdutiofi 
         Caption         =   "&Etiquetas"
         Height          =   525
         Index           =   2
         Left            =   3840
         Picture         =   "frmAreaRecibo.frx":083A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Genera etiquetas"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdOfertas 
         Caption         =   "Estrellas tdas."
         Height          =   400
         Left            =   3960
         TabIndex        =   56
         ToolTipText     =   "Vista preeliminar de las estrellas para enviar a tiendas"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdCamComp 
         Caption         =   "Camb. X comp."
         Height          =   400
         Left            =   2760
         TabIndex        =   55
         ToolTipText     =   "Cambios por comprador en un rango de fechas"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton CmdGenmdb 
         Caption         =   "&Exporta BD"
         Height          =   400
         Left            =   5160
         TabIndex        =   54
         ToolTipText     =   "Exporta Base de datos a Microsoft Access"
         Top             =   1200
         Width           =   1200
      End
      Begin MSAdodcLib.Adodc AdoProd 
         Height          =   330
         Left            =   0
         Top             =   0
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
         Caption         =   "Adoprod"
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
      Begin VB.Label lbltransO 
         Alignment       =   2  'Center
         Caption         =   "Realizando Transacción ......"
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   3840
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.Label avance 
         Height          =   255
         Left            =   5400
         TabIndex        =   71
         Top             =   3840
         Width           =   615
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
      Height          =   1695
      Left            =   3720
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblContra 
         BackColor       =   &H00C0C000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Fraimporta 
      Height          =   5175
      Left            =   1560
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame fracambios 
         Height          =   2175
         Left            =   4080
         TabIndex        =   46
         Top             =   720
         Visible         =   0   'False
         Width           =   4600
         Begin VB.CommandButton cmdcamace 
            Caption         =   "&Procesar"
            Height          =   315
            Left            =   480
            TabIndex        =   50
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdcamreg 
            Caption         =   "&Regresar"
            Height          =   315
            Left            =   2160
            TabIndex        =   51
            Top             =   1680
            Width           =   1095
         End
         Begin VB.OptionButton optcamb 
            Caption         =   "Actualizar las dos areas"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton optcamb 
            Caption         =   "Actualizar area de bodega"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   49
            Top             =   1200
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optcamb 
            Caption         =   "&Actualizar area de piso"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   2655
         End
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Rpt CortIn&v."
         Height          =   500
         Index           =   16
         Left            =   2400
         Picture         =   "frmAreaRecibo.frx":093C
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Catálogo de tiendas y franquicias"
         Top             =   2040
         Width           =   1005
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Corte In&v."
         Height          =   500
         Index           =   15
         Left            =   1320
         Picture         =   "frmAreaRecibo.frx":0E6E
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Corte de Inv. Diario"
         Top             =   2040
         Width           =   1005
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Créditos"
         Height          =   500
         Index           =   14
         Left            =   240
         Picture         =   "frmAreaRecibo.frx":13A0
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Genera informacion de créditos otorgados"
         Top             =   2040
         Width           =   1005
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Entradas"
         Height          =   500
         Index           =   4
         Left            =   4200
         Picture         =   "frmAreaRecibo.frx":18D2
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Actualiza pedidos recibidos en Carbonera"
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Frame fraexporta 
         Height          =   790
         Left            =   1560
         TabIndex        =   74
         Top             =   4320
         Visible         =   0   'False
         Width           =   5775
         Begin VB.CommandButton btnexporta 
            Caption         =   "Procesar"
            Height          =   495
            Left            =   3960
            Picture         =   "frmAreaRecibo.frx":1C14
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   240
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker fecfin 
            Height          =   375
            Left            =   2160
            TabIndex        =   75
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20905985
            CurrentDate     =   37110
         End
         Begin MSComCtl2.DTPicker fecini 
            Height          =   375
            Left            =   120
            TabIndex        =   76
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20905985
            CurrentDate     =   37110
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha de Inicio                           Fecha Final"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   120
            Width           =   3855
         End
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Llegadas "
         Height          =   500
         Index           =   13
         Left            =   5400
         Picture         =   "frmAreaRecibo.frx":1F56
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Actualiza pedidos recibidos en Carbonera"
         Top             =   1560
         Width           =   1005
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Traslados"
         Height          =   500
         Index           =   12
         Left            =   7680
         Picture         =   "frmAreaRecibo.frx":2298
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Cortes"
         Height          =   500
         Index           =   10
         Left            =   6600
         Picture         =   "frmAreaRecibo.frx":25DA
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1560
         Width           =   1000
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Cortes"
         Height          =   500
         Index           =   9
         Left            =   240
         Picture         =   "frmAreaRecibo.frx":291C
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Actualiza precios con los cambios enviados por Oficinas centrales"
         Top             =   1440
         Width           =   1000
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Inventario"
         Height          =   500
         Index           =   8
         Left            =   5400
         Picture         =   "frmAreaRecibo.frx":2C5E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Inventario"
         Height          =   500
         Index           =   11
         Left            =   2400
         Picture         =   "frmAreaRecibo.frx":2FA0
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Genera archivo con el inventario de Bodega"
         Top             =   840
         Width           =   915
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Entradas"
         Height          =   500
         Index           =   7
         Left            =   1320
         Picture         =   "frmAreaRecibo.frx":32E2
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   840
         Width           =   915
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "Salidas"
         Height          =   500
         Index           =   6
         Left            =   240
         Picture         =   "frmAreaRecibo.frx":3624
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   840
         Width           =   1000
      End
      Begin MSComCtl2.Animation ani1 
         Height          =   375
         Left            =   2640
         TabIndex        =   32
         Top             =   2640
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         _Version        =   393216
         FullWidth       =   33
         FullHeight      =   25
      End
      Begin ComctlLib.ProgressBar probar1 
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   3840
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Act.Exi.CDC"
         Height          =   500
         Index           =   5
         Left            =   6600
         Picture         =   "frmAreaRecibo.frx":3966
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Actualiza Existecias de Centro de Distribución Mayoroe"
         Top             =   840
         Width           =   1005
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Cat. Usua."
         Height          =   500
         Index           =   3
         Left            =   1320
         Picture         =   "frmAreaRecibo.frx":3CA8
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Catálogo de Usuarios"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Cat. tiendas"
         Height          =   500
         Index           =   2
         Left            =   2400
         Picture         =   "frmAreaRecibo.frx":3FEA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Catálogo de tiendas y franquicias"
         Top             =   1440
         Width           =   1005
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Act. precios"
         Height          =   500
         Index           =   0
         Left            =   4200
         Picture         =   "frmAreaRecibo.frx":432C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Actualiza precios con los cambios enviados por Oficinas centrales"
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Regresar"
         Height          =   500
         Index           =   1
         Left            =   7680
         Picture         =   "frmAreaRecibo.frx":466E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1560
         Width           =   1000
      End
      Begin MSAdodcLib.Adodc AdoCargos 
         Height          =   330
         Left            =   6240
         Top             =   4680
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
         Caption         =   "Adocargos"
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
      Begin VB.CommandButton Cmdopt 
         Caption         =   "&Despensas"
         Height          =   500
         Index           =   17
         Left            =   4200
         Picture         =   "frmAreaRecibo.frx":47E0
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Actualiza pedidos recibidos en Carbonera"
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblProd 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Productos procesados: 0"
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   3600
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "[RECIBIR]"
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
         Left            =   4080
         TabIndex        =   40
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "[ENVIAR]"
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
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Lbltrans 
         Alignment       =   2  'Center
         Caption         =   "Progreso de la actualizacion"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Visible         =   0   'False
         Width           =   7455
      End
   End
   Begin VB.Frame frabusqueda 
      Caption         =   "Búsqueda de Productos:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   7110
      Left            =   120
      TabIndex        =   26
      Top             =   650
      Visible         =   0   'False
      Width           =   11800
      Begin VB.TextBox txtbusca 
         Height          =   350
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   9255
      End
      Begin VB.Data Daocambio 
         Caption         =   "Daocambio"
         Connect         =   "dBASE III;"
         DatabaseName    =   ""
         DefaultCursorType=   2  'ServerSideCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc AdoMargen 
         Height          =   330
         Left            =   0
         Top             =   4800
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
         Caption         =   "AdoMargen"
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
      Begin VB.CommandButton CmdBusca 
         Caption         =   "&Regresar"
         Height          =   450
         Left            =   10200
         Picture         =   "frmAreaRecibo.frx":4B22
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Regresar a la pantalla principal"
         Top             =   240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid gridbusca 
         Bindings        =   "frmAreaRecibo.frx":4C94
         Height          =   6195
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10927
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         ForeColor       =   64
         HeadLines       =   1.7
         RowHeight       =   17
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "consec"
            Caption         =   "  Clave"
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
            Caption         =   "                    Descripción"
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
            DataField       =   "presentacion"
            Caption         =   "Presentación"
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
            DataField       =   "barraspza"
            Caption         =   "Código Barras"
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
            DataField       =   "Precio1"
            Caption         =   "Pre.Aut."
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
            DataField       =   "ofertado"
            Caption         =   "Oferta"
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
         BeginProperty Column06 
            DataField       =   "activo"
            Caption         =   "Activo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   "dd/MM/yy"
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
         BeginProperty Column07 
            DataField       =   "fecact"
            Caption         =   "Fecha Act"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4680
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   959.811
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodbf 
      Height          =   330
      Left            =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adoprov"
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
      Left            =   3240
      TabIndex        =   33
      Top             =   3500
      Visible         =   0   'False
      Width           =   4935
      Begin ComctlLib.ProgressBar pgb 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   615
      Left            =   120
      TabIndex        =   73
      Top             =   0
      Width           =   11775
      _cx             =   63721762
      _cy             =   63702077
      FlashVars       =   ""
      Movie           =   "C:\PITIC\pitico.swf"
      Src             =   "C:\PITIC\pitico.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   "007888"
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
   End
   Begin VB.Data DaoProdDbf 
      Caption         =   "DaoProdbf"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc AdoProv 
      Height          =   330
      Left            =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adoprov"
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
   Begin MSAdodcLib.Adodc adobus 
      Height          =   375
      Left            =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc Adopreprod 
      Height          =   330
      Left            =   0
      Top             =   4320
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
      Caption         =   "AdoPreprod"
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
   Begin MSAdodcLib.Adodc AdoDescuentos 
      Height          =   330
      Left            =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodescuentos"
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
   Begin MSAdodcLib.Adodc AdoDescprod 
      Height          =   330
      Left            =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "AdoDescprod"
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
   Begin MSAdodcLib.Adodc AdoTfproduc 
      Height          =   375
      Left            =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Adotfproduc"
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
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraMenu 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   525
      Width           =   11750
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   700
         Left            =   7320
         TabIndex        =   45
         Top             =   840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "Sa&lir"
         Height          =   600
         Index           =   11
         Left            =   10800
         Picture         =   "frmAreaRecibo.frx":4CA9
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir del sistema"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "Ofertas"
         Height          =   600
         Index           =   3
         Left            =   9780
         Picture         =   "frmAreaRecibo.frx":5267
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Importar  datos y  catalogos "
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Facturas"
         Height          =   600
         Index           =   10
         Left            =   6915
         Picture         =   "frmAreaRecibo.frx":59E1
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Facturas"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Ventas"
         Height          =   600
         Index           =   9
         Left            =   6000
         Picture         =   "frmAreaRecibo.frx":5B83
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ventas"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Hoja Cat."
         Height          =   600
         Index           =   8
         Left            =   3045
         Picture         =   "frmAreaRecibo.frx":5C9D
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Catalogo X Proveedores"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "Ped. &Aba."
         Height          =   600
         Index           =   7
         Left            =   50
         Picture         =   "frmAreaRecibo.frx":657F
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Pedidos por tienda para abastecimiento"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Utilerias"
         Height          =   600
         Index           =   6
         Left            =   8880
         Picture         =   "frmAreaRecibo.frx":6761
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Utilerias"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Ped. prov"
         Height          =   600
         Index           =   5
         Left            =   1865
         Picture         =   "frmAreaRecibo.frx":68DF
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Pedidos Por Proveedor"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Inventario"
         Height          =   600
         Index           =   4
         Left            =   4875
         Picture         =   "frmAreaRecibo.frx":70A1
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Inventarios"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Reportes"
         Height          =   600
         Index           =   2
         Left            =   7830
         Picture         =   "frmAreaRecibo.frx":722B
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Reportes"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "Ped. &Sug"
         Height          =   600
         Index           =   0
         Left            =   960
         Picture         =   "frmAreaRecibo.frx":772D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Pedidos Sugeridos"
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Traslados"
         Height          =   600
         Index           =   1
         Left            =   3960
         Picture         =   "frmAreaRecibo.frx":790F
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Traslados"
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.PictureBox CR1 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   89
      Top             =   3000
      Width           =   1200
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   29
      Top             =   7875
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmAreaRecibo.frx":7B21
            Text            =   "F2 - PRODUCTO"
            TextSave        =   "F2 - PRODUCTO"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmAreaRecibo.frx":7CA3
            Text            =   "F3 - PRECIOS"
            TextSave        =   "F3 - PRECIOS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmAreaRecibo.frx":7E01
            Text            =   "F4 - BARRAS"
            TextSave        =   "F4 - BARRAS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmAreaRecibo.frx":7F47
            Text            =   "F5 - CLAVE PROD."
            TextSave        =   "F5 - CLAVE PROD."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmAreaRecibo.frx":8059
            Text            =   "F8 - CALCULADORA"
            TextSave        =   "F8 - CALCULADORA"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar stbmensajes 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   30
      Top             =   8175
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   9596
            MinWidth        =   9596
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NÚM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Imgtapiz 
      BorderStyle     =   1  'Fixed Single
      Height          =   6100
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   11775
   End
   Begin VB.Menu mnuorg 
      Caption         =   "&Organizar"
      Begin VB.Menu mnuLin 
         Caption         =   "&Lineas"
      End
      Begin VB.Menu mnuFam 
         Caption         =   "&Familias"
      End
      Begin VB.Menu mnuDep 
         Caption         =   "&Departamento"
      End
      Begin VB.Menu mnuProv 
         Caption         =   "&Proveedores"
      End
   End
   Begin VB.Menu mnupedidos 
      Caption         =   "&Pedidos"
      Begin VB.Menu mnupedprov 
         Caption         =   "Ped. Proveedores"
      End
      Begin VB.Menu mnupedaba 
         Caption         =   "Ped. Abastecimiento"
      End
      Begin VB.Menu MnuPedsug 
         Caption         =   "Ped.&Sug."
      End
   End
   Begin VB.Menu MnuTraslados 
      Caption         =   "&Traslados"
      Begin VB.Menu mnuenvios 
         Caption         =   "Envios"
      End
      Begin VB.Menu mnurecibo 
         Caption         =   "Recibos"
      End
   End
   Begin VB.Menu mnuinventario 
      Caption         =   "&Inventario"
      Begin VB.Menu mnuinvtot 
         Caption         =   "Inv. Total"
      End
      Begin VB.Menu mnuinvbod 
         Caption         =   "Inv. Bodega"
      End
      Begin VB.Menu mnuinvpi 
         Caption         =   "Inv. Piso"
      End
   End
   Begin VB.Menu mnuventas 
      Caption         =   "&Ventas"
      Begin VB.Menu vtamay 
         Caption         =   "Ventas al Mayoreo"
      End
      Begin VB.Menu Mnudespensa 
         Caption         =   "&Despensas"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuutil 
      Caption         =   "&Utilerias"
      Begin VB.Menu mnuactpre 
         Caption         =   "Actualizar Precios"
      End
      Begin VB.Menu mnugeninv 
         Caption         =   "Generar Inventario"
      End
      Begin VB.Menu mnugenent 
         Caption         =   "Generar Entradas"
      End
      Begin VB.Menu mnugensal 
         Caption         =   "Generar Salidas"
      End
      Begin VB.Menu mnugencor 
         Caption         =   "Generar Cortes"
      End
      Begin VB.Menu mnuusu 
         Caption         =   "&Procter"
      End
      Begin VB.Menu verpr 
         Caption         =   "ver Precios"
      End
      Begin VB.Menu mnutdas 
         Caption         =   "Tiendas"
      End
   End
   Begin VB.Menu Mnuacerca 
      Caption         =   "&Acerca de.."
      Begin VB.Menu MnuSolAct 
         Caption         =   "Solo activos"
      End
      Begin VB.Menu MnuRay1 
         Caption         =   "-"
      End
      Begin VB.Menu mnucre 
         Caption         =   "Créditos"
      End
      Begin VB.Menu MnuTapiz 
         Caption         =   "A&signar Tapiz"
      End
   End
   Begin VB.Menu mnuSAlir 
      Caption         =   "&Cerrar"
      Begin VB.Menu MnuSes 
         Caption         =   "Cerrar &sesión"
      End
      Begin VB.Menu MnuCerCon 
         Caption         =   "&Conectar a BD"
      End
      Begin VB.Menu MnuCerRay 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCerR 
         Caption         =   "Salir del programa"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAreaRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StrTeclaPres As String
Private lbus As Boolean
Private nCat
Private nOpUtil As Integer
Dim CNMDB As ADODB.Connection

Private Sub btnexporta_Click()
Lbltrans.Visible = True
probar1.Visible = True
Select Case TIPOEXP
   Case "salida"
        Call salidas
   Case "entrada"
        Call entradas
   Case "inventario"
        Call inventarios
   Case "CORTES"
        Call cortesexp
   Case "corteinv"
        Call rptcorteInv
End Select
fraexporta.Visible = False
End Sub

Private Sub salidas()
'SE PRETENDE GENERAR UN TXT DELIMITADO CON COMAS PARA
'ENVIAR LA INFORMAesCION
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
' CAMBIOS HOY
'men1 = "A continuacion se generara un archivo en donde contiene los traslados enviados..."
'men2 = "Deseas Continuar ? "

RESP = MsgBox(men1 & vbCrLf & men2, vbYesNo, "EXPORTACION")
If RESP = vbNo Then
   Exit Sub
End If
CAD = " SELECT * FROM TRASLADOS, DETALLETRASLADO WHERE t_clave = dt_clave  and t_entrada = 0 and t_enviado = 1 and T_FECHA >=  '" & Me.fecini.Value & "' and t_fecha <= '" & fecfin.Value & "' AND t_motivocancelA IS NULL ORDER BY dt_clave,dt_producto"
rs.Open CAD, cn, adOpenKeyset, adLockReadOnly, adCmdText
'On Error GoTo ERROR:
Open "P:\BUZON\sal" & Mid(cSucursal, 6, 3) & ".TXT" For Output As #1
nreg = 0
probar1.Min = 0
probar1.Max = rs.RecordCount
stbmensajes.Panels(1).Text = "Generando Archivo de Salidas... "
stbmensajes.Refresh
lblProd.Visible = True
While Not rs.EOF
        nreg = nreg + 1
        If probar1.Max < nreg Then
           probar1.Max = probar1.Max + 1000
        End If
        probar1.Value = nreg
        lblProd.Caption = "Productos procesados: " & Str(nreg): lblProd.Refresh
        If rs!t_tipo Then
           tipo = 1
        Else
           tipo = 0
        End If
        If rs!t_papeleria Then
            papeleria = 1
        Else
            papeleria = 0
        End If
        If rs!t_frutas Then
           frutas = 1
        Else
           frutas = 0
        End If
        If rs!t_pan Then
           Pan = 1
        Else
           Pan = 0
        End If
        If rs!t_MERMA Then
           merma = 1
        Else
           merma = 0
        End If
        If rs!t_AJUSTE Then
           ajuste = 1
        Else
           ajuste = 0
        End If
        If rs!t_AUTO Then
           auto = 1
        Else
           auto = 0
        End If
        CAD = Trim(rs!t_clave) & "|" & rs!T_FECHA & "|" & tipo & "|" & rs!t_costo & "|" & rs!t_sucursalemisor & _
              "|" & Trim(rs!t_sucursalreceptor) & "|" & papeleria & "|" & frutas & "|" & Pan & "|" & _
               merma & "|" & auto & "|" & ajuste & "|" & rs!Dt_producto & "|" & _
               rs!dt_cantidad & "|" & rs!dt_cantidadp & "|" & rs!DT_costo & "|" & rs!DT_costop & "|" & rs!DT_IMPORTE & "|" & rs!DT_Iva & "|" & rs!DT_IEPS & "|" & rs!dt_tasaieps & "|" & rs!DT_venta & "|" & rs!DT_ventap & "|" & rs!t_foliotie & "|" & rs!t_pedido
        Print #1, CAD
        rs.MoveNext
Wend
stbmensajes.Panels(1).Text = cMens
stbmensajes.Refresh
MsgBox "Proceso Finalizado...", vbInformation, "EXPORTACION"
probar1.Enabled = False
probar1.Visible = False
lblProd.Visible = False
ani1.AutoPlay = False
Close #1
Exit Sub
Error:
   MsgBox "No existe Informacion O, No existe Conexion de Red...", vbInformation, "EXPORTACION"
End Sub

Private Sub inventarios()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
men1 = "Generacion del Inventario de Bodega "
men2 = "Deseas Continuar ? "
RESP = MsgBox(men1 & vbCrLf & men2, vbYesNo, "EXPORTACION")
If RESP = vbNo Then
   Exit Sub
End If
CAD = " SELECT * FROM Inventario WHERE incant > 0 or incantpza > 0 "
rs.Open CAD, cn, adOpenKeyset, adLockOptimistic, adCmdText
On Error GoTo Error:
ANIMA
Open "P:\BUZON\inv" & Mid(cSucursal, 6, 3) & ".TXT" For Output As #1
'Open "c:\paso\inv" & Mid(cSucursal, 6, 3) & ".TXT" For Output As #1
nreg = 0
probar1.Min = 0
probar1.Max = rs.RecordCount
stbmensajes.Panels(1).Text = "Generando Archivo de Inventario... "
stbmensajes.Refresh
lblProd.Visible = True
While Not rs.EOF
        nreg = nreg + 1
        probar1.Value = nreg
        lblProd.Caption = "Productos procesados: " & Str(nreg): lblProd.Refresh
        pzas = rs!InCantPza
        If IsNull(pzas) Then
           pzas = 0
        End If
        CAD = Trim(rs!Inprod) & "|" & rs!InCant & "|" & pzas
        Print #1, CAD
        rs.MoveNext
Wend
stbmensajes.Panels(1).Text = cMens
stbmensajes.Refresh
MsgBox "Proceso Finalizado...", vbInformation, "EXPORTACION"
Close #1
probar1.Enabled = False
probar1.Visible = False
lblProd.Visible = False
ani1.AutoPlay = False
Exit Sub
Error:
   MsgBox "No existe Informacion O, No existe Conexion de Red...", vbInformation, "EXPORTACION"
End Sub

Private Sub cortesexp()
Dim rs As ADODB.Recordset
men1 = "Proceso para Enviar Información de cortes de caja"
men2 = "Deseas Continuar ? "
RESP = MsgBox(men1 & vbCrLf & men2, vbYesNo + vbQuestion, "EXPORTACION")
If RESP = vbNo Then
   Exit Sub
End If
Set rs = New ADODB.Recordset
CAD = "SELECT * FROM facventa_det WHERE fecha_det >= '" & fecini.Value & "' AND fecha_det <= '" & fecfin.Value & "' AND tasaieps > 0 ORDER BY fecha_det,factura"
rs.Open CAD, cCadConex, adOpenKeyset, adLockReadOnly, adCmdText
On Error GoTo Error:
ANIMA
Open "P:\BUZON\COR" & Mid(cSucursal, 6, 3) & ".TXT" For Output As #1
nreg = 0
probar1.Min = 0
probar1.Max = rs.RecordCount
stbmensajes.Panels(1).Text = "Generando Archivo de Cortes... "
stbmensajes.Refresh
lblProd.Visible = True
While Not rs.EOF
   nreg = nreg + 1
   probar1.Value = nreg
   lblProd.Caption = "Productos procesados: " & Str(nreg): lblProd.Refresh
   CAD = Trim(rs!producto) & "|" & rs!cantidad & "|" & rs!cantidadp & "|" & rs!PRECIO & "|" & rs!preciop & "|" & rs!costo & "|" & rs!costop & "|" & rs!importe & "|" & rs!iva & "|" & rs!ieps & "|" & rs!tasaieps & "|" & _
   Trim(rs!Factura) & "|" & Trim(rs!SERIE) & "|" & "0|" & Trim(rs!rfc_det) & "|" & rs!fecha_det & "|" & Trim(Mid(cSucursal, 1, 3))
   Print #1, CAD
   rs.MoveNext
Wend
stbmensajes.Panels(1).Text = cMens
stbmensajes.Refresh
MsgBox "Proceso Finalizado...", vbInformation, "EXPORTACION"
Close #1
probar1.Enabled = False
probar1.Visible = False
lblProd.Visible = False
ani1.AutoPlay = False
ani1.Visible = False
Exit Sub
Error:
   MsgBox "No existe Informacion O, No existe Conexion de Red...", vbInformation, "EXPORTACION"
   Close #1
End Sub

'SE PRETENDE GENERAR UN TXT DELIMITADO CON COMAS PARA
'ENVIAR LA INFORMACION
Private Sub entradas()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
men1 = "Este proceso Genera Archivo de Entradas de Productos "
men2 = "Deseas Continuar ? "
RESP = MsgBox(men1 & vbCrLf & men2, vbYesNo + vbQuestion, "EXPORTACION")
If RESP = vbNo Then
   Exit Sub
End If
stbmensajes.Panels(1).Text = "Generando archivo de entradas a traves de traslados... "
stbmensajes.Refresh
CAD = " SELECT * FROM TRASLADOS, DETALLETRASLADO WHERE t_clave = dt_clave and t_entrada = 1 and T_FECHA >=  '" & Me.fecini.Value & "' and t_fecha <= '" & fecfin.Value & "'"
rs.Open CAD, cn, adOpenKeyset, adLockReadOnly, adCmdText
'On Error GoTo ERROR:
Open "P:\BUZON\ent" & Mid(cSucursal, 6, 3) & ".TXT" For Output As #1
nreg = 0
probar1.Min = 0
probar1.Max = IIf(rs.RecordCount > 0, rs.RecordCount, 1)
lblProd.Visible = True
While Not rs.EOF
        nreg = nreg + 1
        probar1.Value = nreg
        lblProd.Caption = "Productos procesados: " & Str(nreg): lblProd.Refresh
        If rs!t_tipo Then
           tipo = 1
        Else
           tipo = 0
        End If
        If rs!t_papeleria Then
            papeleria = 1
        Else
            papeleria = 0
        End If
        If rs!t_frutas Then
           frutas = 1
        Else
           frutas = 0
        End If
        If rs!t_pan Then
           Pan = 1
        Else
           Pan = 0
        End If
        If rs!t_MERMA Then
           merma = 1
        Else
           merma = 0
        End If
        If rs!t_AJUSTE Then
           ajuste = 1
        Else
           ajuste = 0
        End If
        If rs!t_AUTO Then
           auto = 1
        Else
           auto = 0
        End If
        CAD = Trim(rs!t_clave) & "|" & rs!T_FECHA & "|" & tipo & "|" & rs!t_costo & "|" & rs!t_sucursalemisor & _
              "|" & rs!t_sucursalreceptor & "|" & papeleria & "|" & frutas & "|" & Pan & _
              "|" & merma & "|" & auto & "|" & ajuste & "|" & rs!Dt_producto & "|" & _
              rs!dt_cantidad & "|" & rs!dt_cantidadp & "|" & rs!DT_costo & "|" & rs!DT_costop & "|" & rs!DT_IMPORTE & "|" & rs!DT_venta & "|" & rs!DT_ventap & "|" & rs!DT_Iva & "|" & rs!DT_IEPS & "|"
        Print #1, CAD
        rs.MoveNext
Wend
stbmensajes.Panels(1).Text = "Generando archivo de entradas a traves de pedidos... "
stbmensajes.Refresh
probar1.Value = 1
Print #1, "P"   'Delimitador para pedidos instantáneos
rs.Close
CAD = "SELECT * FROM pedidos,detallefactura WHERE p_pedido = df_pedido AND df_sugerido = 0 AND p_fecentreal >= '" & Me.fecini.Value & "' and p_fecentreal <= '" & fecfin.Value & "' and p_cancelado = 0 and p_recibido = 1"
rs.Open CAD, cn, adOpenKeyset, adLockReadOnly, adCmdText
probar1.Min = 0
probar1.Max = rs.RecordCount
nreg = 0
While Not rs.EOF
    nreg = nreg + 1
    probar1.Value = nreg
    CAD = Trim(p_Pedido) & "|" & rs!p_proveedor & "|" & rs!p_fecped & "|" & rs!p_sucursal & "|" & rs!p_fecentreal & "|" & rs!df_prod & "|" & rs!df_cantsol & "|" & rs!df_cantreal & "|" & rs!df_cantsolp & "|" & rs!df_cantrealP & "|" & rs!df_Costo & "|"
    Print #1, CAD
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
stbmensajes.Panels(1).Text = cMens
stbmensajes.Refresh
MsgBox "Proceso Finalizado...", vbInformation, "EXPORTACION"
probar1.Enabled = False
probar1.Visible = False
lblProd.Visible = False
ani1.AutoPlay = False
Close #1
Exit Sub
Error:
   MsgBox "No existe Informacion O, No existe Conexion de Red...", vbInformation, "EXPORTACION"
   
End Sub


Private Sub cmdBotones_Click(Index As Integer)
Select Case Index
    Case 0  ' Pedidos
         MnuPed_Click
         'Form1.Show
    Case 1  ' Traslados
         frmtraslados.Show
    Case 2  ' Reportes
         stbmensajes.Panels(1).Text = "Obteniendo informacion necesaria para generar reportes"
         stbmensajes.Refresh
         FrmReport.Show
    Case 3  ' Ofertas
         cMens = frmAreaRecibo.stbmensajes.Panels(1).Text
         frmAreaRecibo.stbmensajes.Panels(1).Text = "Espere un momento, buscando productos ofertados....."
         Fofertasnew.Show
         frmAreaRecibo.stbmensajes.Panels(1).Text = cMens
    Case 4  ' Inventario
         nCat = 0
         nOp = 1 'Solo cuando es 31 no se cargar el frmarearecibo
         stbmensajes.Panels(1).Text = "Espere un momento obteniendo existencia de productos"
         stbmensajes.Refresh
         'CLAVEINVENTARIO = Mid(cSucursal, 1, 3)
         'If CLAVEINVENTARIO = 10 Then
         '   If MsgBox("DESEAS VER SOLAMENTE EL INVENTARIO DE LA BODEGA", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
               frmModInv.Show
         '   Else
         '      frminvtodo.Show
         '   End If
         'Else
             frmModInv.Show
         'End If
    Case 5
         frmpedBod.Show
    Case 6
         fecini.Value = date: fecfin.Value = date
         Cmdopt(0).Enabled = tipotienda <> 1
         If MsgBox("Deseas ver utilerías del área de costos?", vbQuestion + vbYesNo) = vbYes Then
            fraUtiOfi.Visible = True
         Else
            Fraimporta.Visible = True
         End If
    Case 7
          frmpedAba.Show
    Case 8
         fhojacat.Show
    Case 9
         frmPrincipal.Show
         Unload Me
    Case 10
         frmFacturas.Show
         Unload Me
    Case 11
        End
End Select
End Sub

Private Sub cmdcamace_Click()
'UnZip "\\" & SERVIDOR & "\disco-c\paso\HOY.ZIP", "\\" & SERVIDOR & "\disco-c\paso"
If optcamb(0).Value Then
   ActPiso
   ActBodega
ElseIf optcamb(1).Value Then
   ActPiso
ElseIf optcamb(2).Value Then
   ActBodega
End If
ActProv    'Actualizacion de proveedores
MsgBox "LA ACTUALIZACION DEPRECIOS SE REALIZO CORRECTAMENTE", vbInformation
End Sub

Private Sub cmdCamComp_Click()
If Not VERFEC Then Exit Sub
cMens = stbmensajes.SimpleText
Me.stbmensajes.Panels(1).Text = "Espere un momento generando reporte......"
stbmensajes.Refresh
CR1.WindowTitle = "Reporte de cambios por comprador"
CR1.Connect = cCadConex
CR1.Formulas(0) = "ENCA = 'LISTADO DE CAMBIOS DE PRECIO DEL " & txtinicio.Text & "AL " & txtfinal.Text & "'"
CR1.Formulas(1) = "FORMSELEC = {PREPROD.FECHAACT} >= Date (" & Format(txtinicio.Text, "YYYY,MM,DD") & ")  AND {PREPROD.FECHAACT} <= Date (" & Format(txtfinal.Text, "YYYY,MM,DD") & ")"
CR1.ReportFileName = App.Path & "\cambCom.rpt"
CR1.DataFiles(0) = ""
Me.CR1.Action = 1
Me.stbmensajes.Panels(1).Text = cMens
stbmensajes.Refresh
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub cmdcamreg_Click()
lblProd.Visible = False
Me.fracambios.Visible = False
End Sub

Private Sub cmdConAceptar_Click()
If txtContra.Text <> "SAP2004" Then
   If nCat = 3 Then
       'solo gondoleros
       lpprod = True
       frmCatusu.Show 1
       Me.fraCon.Visible = False
   Else
        MsgBox "LA CONTRASEÑA DE ACCESO ES INCORRECTA", vbExclamation
        txtContra.SetFocus
        SendKeys "+{HOME}"
        Exit Sub
   End If
Else
  fraCon.Visible = False
  If nCat = 2 Then
     frmCatTienda.Show 1
  ElseIf nCat = 3 Then
     frmCatusu.Show 1
  Else
    stbmensajes.Panels(1).Text = Space(10) + "Espere un momento cargando inventario de productos"
    stbmensajes.Refresh
    frmModInv.Show
  End If
End If
End Sub

Private Sub cmdConCance_Click()
fraCon.Visible = False
If nCat = 0 Then
    If MsgBox("DESEAS EXCLUSIVAMENTE CONSULTAR INVENTARIO", vbInformation + vbYesNo) = vbYes Then
        stbmensajes.Panels(1).Text = Space(10) + "Espere un momento cargando inventario de productos"
        stbmensajes.Refresh
        frmModInv.Show
        frmModInv.CmdActualizar.Visible = False
    End If
End If
End Sub

Private Sub cmdGenerar_Click()
  actcatprev    'Actualiza los catalogos para preventistas (Handheld)
End Sub

Private Sub CmdGenmdb_Click()
Dim cmen
Dim conexmdb As String
Dim sNameOrig As String
Dim sNameDest As String
Dim vEscenario As Variant
Dim vContra As Variant

     Dim Resultado As Long
     Dim intContadorFicheros As Integer
     Dim FuncionesZip As ZIPUSERFUNCTIONS
     Dim OpcionesZip As ZPOPT
     Dim NombresFicherosZip As ZIPnames
         
     cmen = Me.stbmensajes.Panels(1).Text
     Dim rs As ADODB.Recordset
     Dim RSTEMP As ADODB.Recordset
     Set CNMDB = New ADODB.Connection
     cvesuc = Trim(Mid(cSucursal, 1, 3))
     If Sql Then
        If MsgBox("DESEAS EXPORTAR LA BASE DE DATOS?", vbYesNo + vbQuestion, "Confirma exportación") = vbNo Then Exit Sub
        
        RESP = MsgBox("HACIA DONDE DESEAS ENVIAR LA EXPORTACION DE LA BASE DE DATOS; A LA RED INTERNA?" & Chr(13) & Chr(13) & "[SI ]  Se envía a la red interna de " & Mid(cSucursal, 3) & Chr(13) & "[NO] Se envía al disco duro de esta máquina (Viajes)", vbQuestion + vbYesNo)
        If RESP = vbYes Then
           conexmdb = "DSN=PITICOMDB;DBQ=P:\PITICO\PITICO" & cvesuc & "\PITICO" & cvesuc & ".mdb;DefaultDir=P:\PITICO\PITICO" & cvesuc & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"
        Else
           conexmdb = "DSN=PITICOMDB;DBQ=" & App.Path & "\PITICO" & cvesuc & ".mdb;DefaultDir=" & App.Path & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"
        End If
     Else
        MsgBox "ES ESTA MODALIDAD DE PROGRAMA NO ES POSIBLE EXPORTAR LA BASE DE DATOS", vbCritical
        Exit Sub
     End If
     CNMDB.Open conexmdb
     
     Lbltrans.Visible = True
     Set rs = New ADODB.Recordset
     rs.CursorType = adOpenStatic
     lbltransO.Visible = True
     
     'TABLA DE CARGOS
     stbmensajes.Panels(1).Text = "Exportando catálogo de cargos..."
     stbmensajes.Refresh
     rs.Open "SELECT * FROM cargos,tfproduc WHERE interno = 0 And caprod = consec AND activo = 1", cn
     Call exptablamdb(rs, "cargos")
      
     'TABLA CATIEPS
     stbmensajes.Panels(1).Text = "Exportando catálogo de IEPS..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM catieps", cn
     Call exptablamdb(rs, "catieps")
     
     'TABLA CATPROV
     stbmensajes.Panels(1).Text = "Exportando catálogo de Proveedores..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM catprov WHERE activo = 1", cn
     Call exptablamdb(rs, "catprov")
          
     'tabla CATREPRE
     stbmensajes.Panels(1).Text = "Exportando catálogo de representantes..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM catrepre", cn
     Call exptablamdb(rs, "catrepre")
     
     'tabla CATTIENDA
     stbmensajes.Panels(1).Text = "Exportando catálogo de tiendas..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM cattienda", cn
     Call exptablamdb(rs, "cattienda")
     
     'tabla DEPARTAMENTO
     stbmensajes.Panels(1).Text = "Exportando catálogo de departamentos..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM departamento", cn
     Call exptablamdb(rs, "departamento")
     
     'tabla DESCPROD
     stbmensajes.Panels(1).Text = "Exportando catálogo de descuentos..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM descprod,tfproduc WHERE interno = 0 And producto = consec And activo = 1", cn
     Call exptablamdb(rs, "descprod")
         
     'TABLA DE DESCUENTOS
     stbmensajes.Panels(1).Text = "Exportando catálogo de descuentos con promociones..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM descuentos,tfproduc WHERE interno = 0 And deprod = consec and activo = 1", cn
     Call exptablamdb(rs, "descuentos")
     
     'TABLA DE FAMILIAS
     stbmensajes.Panels(1).Text = "Exportando catálogo de familias..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM familias", cn
     Call exptablamdb(rs, "familias")
     
     'TABLA inventario
     stbmensajes.Panels(1).Text = "Exportando catálogo de inventario..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM inventario", cn
     Call exptablamdb(rs, "inventario")
     
     'TABLA DE LINEAS
     stbmensajes.Panels(1).Text = "Exportando catálogo de lineas..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM lineas", cn
     Call exptablamdb(rs, "lineas")
     
     'TABLA DE LINPROVE
     stbmensajes.Panels(1).Text = "Exportando catálogo de lineas por proveedor..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM linprove", cn
     Call exptablamdb(rs, "linprove")
     
     'TABLA DE MARGEN
     stbmensajes.Panels(1).Text = "Exportando catálogo de margenes..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM margen,tfproduc WHERE interno = 0 And producto = consec And activo = 1", cn
     Call exptablamdb(rs, "margen")
     
     'TABLA PREPROD
     stbmensajes.Panels(1).Text = "Exportando catálogo de precios..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM preprod,tfproduc WHERE interno = 0 And preclave = consec And Activo = 1", cn
     Call exptablamdb(rs, "preprod")
     
     'TABLA TFPRODUC
     stbmensajes.Panels(1).Text = "Exportando catálogo de productos..."
     stbmensajes.Refresh
     rs.Close
     rs.Open "SELECT * FROM tfproduc WHERE interno = 0 and activo = 1", cn
     Call exptablamdb(rs, "tfproduc")
     stbmensajes.Panels(1).Text = cmen
     stbmensajes.Refresh
     CNMDB.Close
     Set CNMDB = Nothing
     
    If RESP = vbYes Then
     stbmensajes.Panels(1).Text = "Compactando base de datos..."
     stbmensajes.Refresh
     If Dir("P:\PITICO\PITICO16\PITICO16.bak") <> "" Then
        Kill "P:\PITICO\PITICO16\PITICO16.bak"
     End If
     sNameOrig = "P:\PITICO\PITICO16\pitico16.MDB"
     sNameDest = "P:\PITICO\PITICO16\pitico16.BAK"
     vEscenario = dbLangSpanish    ' en el supuesto de BD en Español
     vContra = ";pwd=PORTATIL"   ' Atencion: no olvides el punto y coma
     DBEngine.CompactDatabase sNameOrig, sNameDest, vEscenario, , vContra
     Dim fs As Object
     Set fs = CreateObject("Scripting.FileSystemObject")
     fs.copyfile "P:\PITICO\PITICO16\PITICO16.bak", "P:\PITICO\PITICO16\PITICO16.mdb", True
     Kill "P:\PITICO\PITICO16\PITICO16.bak"
     If RESP = vbNo Then
        fs.copyfile "P:\PITICO\PITICO16\PITICO16.mdb", App.Path & "\PITICO16.mdb", True
     End If
     
     stbmensajes.Panels(1).Text = "Empaquetando base de datos..."
     stbmensajes.Refresh
     FuncionesZip.DLLComment = DevolverDireccionMemoria(AddressOf FuncionParaProcesarComentarios)
     FuncionesZip.DLLPassword = DevolverDireccionMemoria(AddressOf FuncionParaProcesarPassword)
     FuncionesZip.DLLPrnt = DevolverDireccionMemoria(AddressOf FuncionParaProcesarMensajes)
     FuncionesZip.DLLService = DevolverDireccionMemoria(AddressOf FuncionParaProcesarServicios)
     NombresFicherosZip.s(0) = "P:\PITICO\PITICO16\PITICO16.MDB"
     Resultado = ZpInit(FuncionesZip)
     Resultado = ZpSetOptions(OpcionesZip)
     Resultado = ZpArchive(1, "P:\PITICO\PITICO16\PITICO16.Zip", NombresFicherosZip)
    End If
     stbmensajes.Panels(1).Text = cmen
     stbmensajes.Refresh
     CmdutiTer_Click
End Sub

Sub exptablamdb(rs As Recordset, Tabla As String, Optional numero As Integer)
Dim N As Integer
Dim rsttemp As ADODB.Recordset
    numero = IIf(numero = 0, 1, numero)
    probar1o.Min = 0
    CNMDB.Execute "DELETE FROM " & Tabla
    Set rsttemp = New ADODB.Recordset
    rsttemp.Open "SELECT * FROM " & Tabla, CNMDB, adOpenDynamic, adLockOptimistic, adCmdText
    probar1o.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
    c = 0
    campos = rsttemp.Fields.Count
    While Not rs.EOF
        c = c + 1
        lbltransO.Caption = c
        lbltransO.Refresh
        probar1o.Value = c
        rsttemp.AddNew
        For N = 0 To campos - numero
            'MsgBox rsttemp.Fields(N).Name
            rsttemp.Fields(N).Value = rs.Fields(rsttemp.Fields(N).Name).Value
        Next
        rsttemp.Update
        rs.MoveNext
    Wend
    probar1o.Value = 1
    rsttemp.Close
    Set rsttemp = Nothing
End Sub

Private Sub cmdliscam_Click()
On Error GoTo Error:
strform1 = Format(date, "LONG DATE") & "  " & Format(Time, "hh:mm AM/PM")
CR1.WindowTitle = "Cambios a Productos"
CR1.ReportFileName = App.Path & "\CAMBIOSHOY.RPT"
CR1.DataFiles(0) = "P:\paso\cambios.dbf"
CR1.Formulas(0) = "HORA = '" & Trim(strform1) & "'"
CR1.Formulas(1) = ""
CR1.Action = 1
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Cmdlistagte_Click()
If MsgBox("DESEAS VISUALIZAR PRODUCTOS SOLAMENTE CON EXISTENCIA", vbQuestion + vbYesNo) = vbYes Then
   'COND = "{INVENTARIO.incant} > 0 OR {INVENTARIO.incantpza} >0 "
   COND = "{INVENTARIO.incant} > 0 OR {INVENTARIO.incantpza} > 0 "
   encab = "LISTADO DE PRODUCTOS CON EXISTENCIA"
Else
   COND = "{TFPRODUC.ACTIVO} = 1 "
   encab = "LISTADO DE PRODUCTOS ACTIVOS"
End If
CR1.Connect = cn.ConnectionString
CR1.WindowTitle = "Listado de precios para agentes"
CR1.ReportFileName = App.Path & "\listagte.rpt"
CR1.Formulas(0) = "FORMSELEC = " & COND & ""
CR1.Formulas(1) = "ENCAB = '" & encab & "'"
CR1.Action = 1
End Sub

Private Sub cmdOfertas_Click()
Dim ApDoc As Word.Application
Dim rs As ADODB.Recordset
Dim N As Integer
cMensAnt = Me.stbmensajes.Panels(1).Text
Me.stbmensajes.Panels(1).Text = "Espere un momento generando reporte....."
stbmensajes.Refresh
'On Error GoTo Error:
cmdlg.DialogTitle = "Ruta donde se grabará el archivo de Ofertas"
cmdlg.InitDir = "C:\"
cmdlg.Filter = "Archivos Microsoft Word (*.doc) | *.doc"
cmdlg.CancelError = True
cmdlg.ShowSave
Set ApDoc = CreateObject("word.Application")  'run it
ApDoc.Visible = True
ApDoc.Documents.Open FileName:=App.Path & "\noborrar1.doc", ConfirmConversions:=False, _
        ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto
ApDoc.ActiveDocument.Shapes("Picture 5").Select
ApDoc.Selection.Copy
ApDoc.Documents.Close False
With ApDoc
    .Documents.Add DocumentType:=wdNewBlankDocument
    .ActiveDocument.PageSetup.Orientation = wdOrientPortrait
    .ActiveDocument.PageSetup.PaperSize = wdPaperLegal
End With

N = 0
Set rs = New ADODB.Recordset
rs.Open "SELECT descripc, TFPRODUC.medida as medida, tfproduc.contenid as contenid, precio1,precio1ant,preciostemp.fechaini, preciostemp.fechafin FROM PRECIOSTEMP,TFPRODUC WHERE consec = producto ORDER BY descripc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    
ApDoc.Selection.Font.Name = "Bremen Bd BT"
While Not rs.EOF
     descr = IIf(rs!CONTENID < 1, rs!descripc & Str(rs!CONTENID * 1000) & " " & rs!medida, rs!descripc & Str(rs!CONTENID) & " " & rs!medida)
     nPos = Int(Len(descr) / 2)
     nEsp = 1
     Do  'Se imprimen palabras completas
        nComp = InStr(nEsp, descr, " ")
        nEspacio = nComp
        nEsp = nComp + 1
     Loop Until nEspacio >= nPos Or nComp = 0
     nPos = nEspacio - 1
     With ApDoc
        .Selection.EndKey Unit:=wdLine  'Agrego linea
        .Selection.Font.Size = 80
        .Selection.Font.Outline = False
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="OFERTA"
        .Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        .Selection.EndKey Unit:=wdLine  'Agrego linea
        .Selection.TypeParagraph
        .Selection.Font.Size = 34
        .Selection.Font.Shadow = True
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        If nComp = 0 Then
            .Selection.TypeText Text:=descr
        Else
            .Selection.TypeText Text:=Mid(descr, 1, nPos)
            .Selection.TypeParagraph
            .Selection.TypeText Text:=Mid(descr, nPos + 1)
        End If
        .Selection.Font.Shadow = False
        .Selection.TypeParagraph
        .Selection.Font.Size = 85
        .Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        .Selection.EndKey Unit:=wdLine  'Agrego linea
        .Selection.TypeText Text:=Format(rs!precio1, "$ ###,###,#0.00")
        .Selection.EndKey Unit:=wdLine  'Agrego linea
        .Selection.Font.Size = 18
        .Selection.Font.Outline = True
        .Selection.TypeParagraph
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="VIGENCIA DEL " & UCase(Format(DateAdd("d", 1, rs!FechaIni), "dd MMM.")) & " AL " & UCase(Format(rs!FechaFin, "dd MMM."))
        .Selection.TypeParagraph
        .Selection.TypeText Text:="O HASTA AGOTAR EXISTENCIAS"
        .Selection.EndKey Unit:=wdStory  'Fin de seccion 2
        If (N + 1) Mod 2 = 0 And N > 0 Then
            .Selection.InsertBreak Type:=wdPageBreak
        Else
            .Selection.Paste
            .Selection.EndKey Unit:=wdStory
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeParagraph
        End If
        N = N + 1
    End With
    rs.MoveNext
Wend
ActiveDocument.SaveAs FileName:=cmdlg.FileName, FileFormat:=wdFormatDocument _
        , LockComments:=False, Password:="", AddToRecentFiles:=True, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
         SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False
ApDoc.Quit
Set ApDoc = Nothing
Me.stbmensajes.Panels(1).Text = cMensAnt
stbmensajes.Refresh
Exit Sub
Error:
   MsgBox Err.Description, vbCritical
End Sub

Private Sub Cmdopt_Click(Index As Integer)
Dim rs As ADODB.Recordset
Dim mes(11) As Variant
  mes(0) = "ENE": mes(1) = "FEB": mes(2) = "MAR": mes(3) = "ABR":  mes(4) = "MAY": mes(5) = "JUN": mes(6) = "JUL"
  mes(7) = "AGO": mes(8) = "SEP":  mes(9) = "OCT": mes(10) = "NOV": mes(11) = "DIC"
On Error GoTo Error:
Select Case Index
Case 0
   If tipotienda = 1 Then
   '   MsgBox "No es posible Realizar este Proceso en este Tipo de Tienda", vbInformation
   '   Exit Sub
   End If
   Me.fracambios.Visible = True
Case 1
   'For N = 0 To 7
   '    cmdBotones(N).Enabled = True
   'Next
   Fraimporta.Visible = False
   Me.cmdBotones(4).SetFocus
Case 2
   nCat = Index
   fraCon.Visible = True
   txtContra.SetFocus
   txtContra.Text = ""
Case 3
   nCat = Index
   fraCon.Visible = True
   txtContra.SetFocus
   txtContra.Text = ""
Case 4
     Call imppedtie  'Importa pedidos instantaneos o directos recibidos en carbonera
Case 5
     ActExiCdc
Case 6
     fecini.Value = Format(date, "DD/MM/YY")
     fecfin.Value = Format(date, "DD/MM/YY")
     fraexporta.Enabled = True
     fraexporta.Visible = True
     TIPOEXP = "salida"
Case 7
     fecini.Value = Format(date, "DD/MM/YY")
     fecfin.Value = Format(date, "DD/MM/YY")
     fraexporta.Enabled = True
     fraexporta.Visible = True
     TIPOEXP = "entrada"
Case 11
     Call inventarios
Case 8
     Call Inventariosimp
Case 9
     fraexporta.Enabled = True
     fraexporta.Visible = True
     TIPOEXP = "CORTES"
Case 10
     Call cortesimp
Case 12
     Call importaenvios
Case 13
     Call Importallegadas
Case 14
     Dim fs As Object
     Set fs = CreateObject("Scripting.FileSystemObject")
     cArch = "P:\buzon\credito" & Mid(cSucursal, 1, 2) & ".rpt"
     If Dir(cArch) <> "" Then
        Kill cArch
     End If
     CR1.ReportFileName = App.Path & "\facpend.rpt"
     CR1.Destination = crptToFile
     CR1.PrintFileName = cArch
     CR1.PrintFileType = crptCrystal
     CR1.Action = 1
     MsgBox "La información se genero correctamente", vbInformation, "Créditos"
Case 15
    ' Call corteinv    'Corte diario de Inventario
Case 16
     TIPOEXP = "corteinv"
     fraexporta.Visible = True
Case 17
      Call importadesp
End Select
Exit Sub
Error:
   MsgBox Err.Description
End Sub

Private Sub imppedtie()
Dim cArch  As String
MenAnt = stbmensajes.SimpleText
Dim rstbusca As ADODB.Recordset
Set rstbusca = New ADODB.Recordset
cmdlg.FileName = ""
cmdlg.Filter = "Archivos de texto (*.txt) | *.txt"
cmdlg.ShowOpen
cRutArc = cmdlg.FileName
If cRutArc = "" Or IsNull(cRutArc) Then
    MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
    Exit Sub
End If
rutac = cRutArc
For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
Next
cruta = Mid(cRutArc, 1, nPos)
cArch = Mid(cRutArc, nPos + 1)
cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.
'apertura del archivo
Open rutac For Input As #1
If EOF(1) Then
   MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
   Exit Sub
End If
SUC = Pedsuc(cArch)
If IsNull(SUC) Or Trim(SUC) = "" Then
      MsgBox "EL NOMBRE DEL ARCHIVO: " & cArch & "NO ESTA REGISTRADO EN EL SISTEMA", vbCritical, "ENVIOS"
      Exit Sub
End If
Dim Pedido As String
cveant = ""
Lbltrans.Visible = True
tipent = "T"
While Not EOF(1)
   Line Input #1, CAD
   If CAD = "P" Then tipent = "P"
   If tipent = "T" Then
        COSTOT = 0
        'SE LEE EL ARCHIVO TXT
        nreg = nreg + 1
        'probar1.Value = nreg
        Lbltrans = nreg: Lbltrans.Refresh
        lblProd.Caption = "Registros procesados: " & Str(nreg): lblProd.Refresh
        pos1 = InStr(CAD, "|")
        traslado = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        fecha = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        tipo = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        costo = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        emisor = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        receptor = Trim(Mid(CAD, 1, pos1 - 1))
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        papeleria = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        frutas = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        Pan = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        merma = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        auto = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        ajuste = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        producto = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        cantidad = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        cantidadp = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        costod = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        costodp = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        importe = Mid(CAD, 1, pos1 - 1)
       
        If cveant <> traslado Then
            CADENA = "select * from traslados where t_clave = '" & Trim(traslado) & "'"
            rstbusca.Open CADENA, cn, adOpenKeyset, adLockOptimistic, adCmdText
            If rstbusca.EOF Then
                'SE INSERTA EL ENVIO
                perenv = 0: perrec = 0
                enviado = 1: recibido = 0: entrada = 1
                devo = 0
                TRAS = 1
                CADENA = "insert into TRASLADOS (T_CLAVE,T_FECHA,T_TIPO,T_PERENV,T_PERREC,T_COSTO,T_SUCURSALEMISOR,T_SUCURSALRECEPTOR,T_ENVIADO,T_RECIBIDO, T_ENTRADA,T_FOLIOTIE,T_DEVOLUCION,T_PAPELERIA,T_FRUTAS,T_PAN,T_MERMA,T_AJUSTE,t_perfle) VALUES( " & _
                "'" & Trim(traslado) & "','" & fecha & "'," & tipo & "," & perenv & "," & perrec & "," & costo & "," & emisor & "," & receptor & "," & enviado & "," & recibido & "," & entrada & "," & foliotie & "," & devo & "," & papeleria & "," & frutas & "," & Pan & "," & merma & "," & ajuste & ",0)"
                cn.Execute CADENA
            End If
            rstbusca.Close
        End If
        'AHORA SE BUSCA LA PARTE DEL DETALLE POR PRODUCTO
        CADENA = "SELECT * FROM detalletraslado WHERE dt_clave = '" & Trim(traslado) & "' and dt_producto = '" & Trim(producto) & "'"
        rstbusca.Open CADENA, cn, adOpenKeyset, adLockOptimistic, adCmdText
        If rstbusca.EOF Then
            costop = 0
            'SE INSERTA EL detalle del envio
            CADENA = "INSERT INTO detalletraslado(dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido,dt_costo,dt_costop,dt_importe,dt_venta,dt_ventap,dt_iva,dt_ieps) values(" & _
                "'" & Trim(traslado) & "','" & Trim(producto) & "'," & cantidad & "," & cantidadp & ",'0'," & costod & "," & costodp & "," & importe & "," & venta & "," & ventaP & "," & iva & "," & ieps & ")"
            cn.Execute CADENA
        Else
            'cn.Execute "UPDATE detalletraslado set dt_cantidad =" & dt_cantidadp
        End If
        rstbusca.Close
    Else
        'asi se graba en el TXT
        'cad = Trim(p_pedido) & "|" & rs!p_proveedor & "|" & rs!p_fecped & "|" & rs!p_sucursal & "|" & rs!p_fecentreal & "|" & rs!df_prod & "|" & rs!df_cantsol & "|" & rs!df_cantreal & "|" & rs!df_cantsolp & "|" & rs!df_cantrealP & "|" & rs!df_Costo & "|"
        nreg = nreg + 1
        'probar1.Value = nreg
        Lbltrans = nreg: Lbltrans.Refresh
        lblProd.Caption = "Registros procesados: " & Str(nreg): lblProd.Refresh
        pos1 = InStr(CAD, "|")
        Pedido = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        PROVEEDOR = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        fecelab = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        sucursal = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        fecreal = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        producto = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        cantsol = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        cantreal = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        cantsolp = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        cantrealp = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        costo = Mid(CAD, 1, pos1 - 1)
        If cveant <> traslado Then
            CADENA = "SELECT * FROM pedidos WHERE p_pedido = '" & Trim(traslado) & "'"
            rstbusca.Open CADENA, cn, adOpenKeyset, adLockOptimistic, adCmdText
            If rstbusca.EOF Then
                'SE INSERTA EL ENVIO
                enviado = 1: recibido = 0: entrada = 1
                devo = 0
                TRAS = 1
                CADENA = "INSERT INTO detallefactura (p_pedido,p_proveedor,) VALUES( " & _
                "'" & Trim(traslado) & "','" & fecha & "'," & tipo & "," & perenv & "," & perrec & "," & costo & "," & emisor & "," & receptor & "," & enviado & "," & recibido & "," & entrada & "," & foliotie & "," & devo & "," & papeleria & "," & frutas & "," & Pan & "," & merma & "," & ajuste & ",0)"
                cn.Execute CADENA
            End If
            rstbusca.Close
        End If

    
    End If
    cveant = traslado
Wend
End Sub


Private Sub Importallegadas()
'Importar llegadas de pedidos por proveedor de Bodega Carbonera
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim rstBus As ADODB.Recordset
Dim cArch  As String
Dim cProv As String
Dim SUC As String
Dim cnFoxPro As ADODB.Connection
Dim lExiste  As Boolean
Dim EsBack As Boolean
On Error GoTo Error:
   MenAnt = stb1.SimpleText
   cmdlg.DialogTitle = "Abrir archivo enviado por Bodega Carbonera"
   cmdlg.FileName = ""
   cmdlg.CancelError = True   'Para que se genere error al hacer click en el boton cancelar
   cmdlg.Filter = "Archivos Visual Foxpro (*.dbf) | *.dbf"
   cmdlg.ShowOpen
   cRutArc = cmdlg.FileName
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
   Adodbf.CursorType = adOpenKeyset
   Adodbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
   Adodbf.RecordSource = "SELECT * FROM " & cArch
   Adodbf.Refresh
   If Adodbf.Recordset.BOF And Adodbf.Recordset.EOF Then
      MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
      Exit Sub
   'CUANDO SE IMPORTO PREVIEMENTE, SE MARCA EL PRIMER REGISTRO COMO IMPORTADO
   ElseIf Adodbf.Recordset!Importado Then
      MsgBox "EL ARCHIVO SELECCIONADO YA FUE IMPORTADO", vbInformation
      Exit Sub
   End If
   nreg = 0
   probar1.Min = 0
   probar1.Max = Adodbf.Recordset.RecordCount
   stbmensajes.Panels(1).Text = "Generando Archivo de Salidas... "
   stbmensajes.Refresh
   lblProd.Visible = True
   Open App.Path & "\LLEGACAR.TXT" For Output As #1   ' Abre el archivo para operaciones de salida.
   Print #1, Tab(20); "VIVERES Y LICORES S.A DE C.V."
   Print #1, Tab(10); "CARBONERA 1016 COL. TRINIDAD DE LAS HUERTAS"
   Print #1, "LLEGADAS DE PEDIDOS DE CARBONERA RECIBIDOS EL "; UCase(Format(date, "long date"))
   Print #1,   ' Imprime una línea en blanco en el archivo.
   Print #1, "=========================================================================================="
   Print #1, "PEDIDO        FEC. ELAB.        FEC.RECEPCION         PROVEEDOR"
   Print #1, "=========================================================================================="
   
   Set rstBus = New ADODB.Recordset  'Sirve para saber si existe o no el pedido
   Set rsttemp = New ADODB.Recordset
   Set rstins = New ADODB.Recordset
   rstBus.ActiveConnection = cn
   cNotAnt = "": nPed = 0
   Adodbf.Recordset.MoveFirst
   While Not Adodbf.Recordset.EOF
           'Verifico si existe el detalle de la nota
           nreg = nreg + 1
           probar1.Value = nreg
           lblProd.Caption = "Productos procesados: " & Str(nreg): lblProd.Refresh
           If cNotAnt <> Adodbf.Recordset!Clavenota Then
              'SE TIENE QUE CHECAR QUE EL PEDIDO GLOBAL EXISTE, TALVEZ SE GENERO EN LA BODEGA
              rstBus.Open "SELECT * FROM pedprove WHERE pp_pedido  = '" & Mid(Adodbf.Recordset!Clavenota, 2) & "'"
              If Not rstBus.EOF Then
                 pr = InStr(Adodbf.Recordset!Clavenota, "-")
                 PRO = Mid(Adodbf.Recordset!Clavenota, pr - 3, 3)
                 cn.Execute "UPDATE PEDPROVE SET PP_recibe = 1, Pp_FecRecibe = '" & Adodbf.Recordset!FecRec & "',PP_notent = '" & Adodbf.Recordset!Clavenota & "' WHERE PP_PEDIDO = '" & Mid(Adodbf.Recordset!Clavenota, 2) & "'"
                 cFecPed = rstBus!PP_FECHAGEN: cFecConf = rstBus!pp_fecconfirma
              Else
                 pr = InStr(Adodbf.Recordset!Clavenota, "-")
                 PRO = Mid(Adodbf.Recordset!Clavenota, pr - 3, 3)
                 EsBack = Len(Mid(Adodbf.Recordset!Clavenota, 1, pr)) >= 6
                 If EsBack Then
                    CADENA = "INSERT INTO PEDPROVE(pp_proveedor,pp_pedido,pp_fechagen,pp_fecrecibe,pp_notent,pp_observa, pp_pedback, pp_recibe) values(" & _
                             "'" & PRO & "','" & Mid(Adodbf.Recordset!Clavenota, 2) & "','" & Adodbf.Recordset!FecRec & "','" & Adodbf.Recordset!FecRec & "','" & Adodbf.Recordset!Clavenota & "','PEDIDO GENERADO EN CARBONERA','" & Mid(Adodbf.Recordset!Clavenota, pr - 3) & "',1)"
                 Else
                    CADENA = "INSERT INTO PEDPROVE(pp_proveedor,pp_pedido,pp_fechagen,pp_fecrecibe,pp_notent,pp_observa,pp_recibe) values(" & _
                            "'" & PRO & "','" & Mid(Adodbf.Recordset!Clavenota, 2) & "','" & Adodbf.Recordset!FecRec & "','" & Adodbf.Recordset!FecRec & "','" & Adodbf.Recordset!Clavenota & "','PEDIDO GENERADO EN CARBONERA',1)"
                 End If
                 cn.Execute CADENA
                 cFecPed = Space(8): cFecConf = Space(8)
                 'Verifico si es backorder
                 If EsBack Then
                    'Agrego las facturas del backorder al pedido al que pertenecen
                    rsttemp.Open "SELECT * FROM NOTAENTRADA WHERE CLAVENOTA = 'N" & Mid(Adodbf.Recordset!Clavenota, pr - 3) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
                    If Not rsttemp.EOF Then
                    For f = 1 To 10
                        Factura = "Factura" & Trim(Str(f))
                        impfact = "ImpFac" & Trim(Str(f))
                        If Not IsNull(Adodbf.Recordset.Fields(Factura).Value) Then
                           lInd = False
                           For p = 1 To 10
                               FACNOTENT = "Factura" & Trim(Str(p))
                               IMPNOTENT = "ImpFac" & Trim(Str(p))
                               If rsttemp.Fields(IMPNOTENT).Value = 0 Or Trim(rsttemp.Fields(FACNOTENT).Value) = "" Or IsNull(rsttemp.Fields(FACNOTENT).Value) Then
                                  rsttemp.Fields(FACNOTENT).Value = Adodbf.Recordset(Factura).Value
                                  rsttemp.Fields(IMPNOTENT).Value = Adodbf.Recordset(impfact).Value
                                  rsttemp.Update
                                  lInd = True
                                  Exit For
                               End If
                           Next
                           If Not lInd Then
                              MsgBox "ERROR CRITICO; EL NUMERO DE FACTURAS DE TODOS LOS BACKORDERS Y EL PEDIDO EXCEDEN DE 10 FACTURAS, INFORME AL ADMINISTRADOR DEL SISTEMA PARA QUE SE INCREMENTEN CAMPOS YA QUE NO SE GUARDARAN LAS FACTURAS CON SUS IMPORTES CORRESPONDIENTES", vbCritical
                              MsgBox "FOLIO DEL PEDIDO " & Adodbf.Recordset!Clavenota
                           End If
                        End If
                    Next
                    End If
                    rsttemp.Close
                 End If
              End If
              rstBus.Close
              
              rstBus.Open "SELECT * FROM catprov WHERE prove = '" & PRO & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
              Print #1, Mid(Adodbf.Recordset!Clavenota, 2); cFecPed; Space(22 - Len(cFecPed)); Adodbf.Recordset!FecRec; Space(22 - Len(Adodbf.Recordset!FecRec)); PRO; IIf(rstBus.RecordCount = 0, "", "  " & Mid(rstBus!NOMPROVE, 1, 38))
              rstBus.Close
              
              'se debe buscar el maestro de la nota
              CADENA = "SELECT * FROM notaentrada WHERE clavenota =  '" & Adodbf.Recordset!Clavenota & "'"
              rstBus.Open CADENA
              If rstBus.EOF Then
                    'ES NECESARIO INSERTAR LA NOTA PARA QUE SE ACCESE AL DETALLE
                    'CHECAR QUE NO TENGA NULOS
                    FAC1 = IIf(Not IsNull(Adodbf.Recordset!factura1), Adodbf.Recordset!factura1, "0")
                    FAC2 = IIf(Not IsNull(Adodbf.Recordset!factura2), Adodbf.Recordset!factura2, "0")
                    fac3 = IIf(Not IsNull(Adodbf.Recordset!Factura3), Adodbf.Recordset!Factura3, "0")
                    fac4 = IIf(Not IsNull(Adodbf.Recordset!factura4), Adodbf.Recordset!factura4, "0")
                    fac5 = IIf(Not IsNull(Adodbf.Recordset!factura5), Adodbf.Recordset!factura5, "0")
                    fac6 = IIf(Not IsNull(Adodbf.Recordset!factura6), Adodbf.Recordset!factura6, "0")
                    fac7 = IIf(Not IsNull(Adodbf.Recordset!factura7), Adodbf.Recordset!factura7, "0")
                    fac8 = IIf(Not IsNull(Adodbf.Recordset!factura8), Adodbf.Recordset!factura8, "0")
                    fac9 = IIf(Not IsNull(Adodbf.Recordset!factura9), Adodbf.Recordset!factura9, "0")
                    fac10 = IIf(Not IsNull(Adodbf.Recordset!factura10), Adodbf.Recordset!factura10, "0")
                    'IMPORTES
                    IMP1 = IIf(Not IsNull(Adodbf.Recordset!Impfac1), Adodbf.Recordset!Impfac1, "0")
                    imp2 = IIf(Not IsNull(Adodbf.Recordset!Impfac2), Adodbf.Recordset!Impfac2, "0")
                    imp3 = IIf(Not IsNull(Adodbf.Recordset!Impfac3), Adodbf.Recordset!Impfac3, "0")
                    imp4 = IIf(Not IsNull(Adodbf.Recordset!Impfac4), Adodbf.Recordset!Impfac4, "0")
                    Imp5 = IIf(Not IsNull(Adodbf.Recordset!Impfac5), Adodbf.Recordset!Impfac5, "0")
                    imp6 = IIf(Not IsNull(Adodbf.Recordset!Impfac6), Adodbf.Recordset!Impfac6, "0")
                    imp7 = IIf(Not IsNull(Adodbf.Recordset!Impfac7), Adodbf.Recordset!Impfac7, "0")
                    imp8 = IIf(Not IsNull(Adodbf.Recordset!Impfac8), Adodbf.Recordset!Impfac8, "0")
                    imp9 = IIf(Not IsNull(Adodbf.Recordset!Impfac9), Adodbf.Recordset!Impfac9, "0")
                    imp10 = IIf(Not IsNull(Adodbf.Recordset!Impfac10), Adodbf.Recordset!Impfac10, "0")
                    CADENA = "INSERT INTO NOTAENTRADA(PEDIDO,CLAVENOTA,factura1,impfac1,factura2,impfac2,factura3,impfac3,factura4,impfac4,factura5,impfac5,factura6,impfac6,factura7,impfac7,factura8,impfac8,factura9,impfac9,factura10,impfac10) VALUES('" & _
                    Mid(Adodbf.Recordset!Clavenota, 2) & "','" & Adodbf.Recordset!Clavenota & "','" & Trim(FAC1) & "'," & IMP1 & ",'" & Trim(FAC2) & "'," & imp2 & ",'" & fac3 & "'," & imp3 & _
                    ",'" & fac4 & "'," & imp4 & ",'" & fac5 & "'," & Imp5 & ",'" & fac6 & "'," & imp6 & ",'" & fac7 & "'," & imp7 & _
                    ",'" & fac8 & "'," & imp8 & ",'" & fac9 & "'," & imp9 & ",'" & fac10 & "'," & imp10 & ")"
                    'MsgBox cadena
                    cn.Execute CADENA
              Else
                  'CRITERIO PARA QUE HACER EN EL CASO DE QUE YA EXISTA, SUPUESTAMENTE
                  'NO SE DEBE REPETIR EL ENVIO
              End If
              rstBus.Close
              nPed = nPed + 1
           End If
           'EN EL CASO DE QUE EL PEDIDO NO HAYA EXISTIDO , TAMBIEN SE DEBEN AGREGAR
           ' LOS PRODUCTO EN EL DETALLEGLOBAL
           CADENA = "SELECT * FROM DETALLEGLOBAL WHERE DG_PEDIDO = '" & Mid(Adodbf.Recordset!Clavenota, 2) & "' AND DG_PRODUCTO = '" & Trim(Adodbf.Recordset!producto) & "'"
           rstBus.Open CADENA
           If Not rstBus.EOF Then
              CADENA = "UPDATE DETALLEglobal SET dg_cantreal = " & Adodbf.Recordset!CantRecC & ", dg_cantsol = " & Adodbf.Recordset!CantSolc & ",dg_promocionr = " & Adodbf.Recordset!promrec & ",dg_costo = " & Adodbf.Recordset!costo & " where DG_PEDIDO = '" & Mid(Adodbf.Recordset!Clavenota, 2) & "' and dg_producto = '" & Trim(Adodbf.Recordset!producto) & "'"
              cn.Execute CADENA
           Else
              'Si es BackOrder se acumula al pedido que corresponde
              If EsBack Then
                 cn.Execute "UPDATE detalleglobal SET dg_cantreal = dg_cantreal + " & Adodbf.Recordset!CantRecC & ", dg_promocionr = dg_promocionr + " & Adodbf.Recordset!promrec & " WHERE dg_pedido = '" & Mid(Adodbf.Recordset!Clavenota, pr - 3) & "' AND dg_producto = '" & Trim(Adodbf.Recordset!producto) & "'"
              End If
              'Ahora agrego el producto al detalleglobal
              CADENA = "INSERT INTO detalleglobal (dg_cantreal,dg_cantsol,dg_promocion,dg_producto,dg_pedido,dg_promocionr,dg_costo) values(" & _
                        Adodbf.Recordset!CantRecC & "," & Adodbf.Recordset!CantSolc & "," & Adodbf.Recordset!promrec & ",'" & Trim(Adodbf.Recordset!producto) & "','" & Mid(Adodbf.Recordset!Clavenota, 2) & "'," & Adodbf.Recordset!promrec & "," & Adodbf.Recordset!costo & ")"
              cn.Execute CADENA
           End If
           rstBus.Close
           CADENA = "SELECT * FROM detallenota WHERE producto = '" & Adodbf.Recordset!producto & "' AND ClaveNota = '" & Adodbf.Recordset!Clavenota & "'"
           rstBus.Open CADENA
           lExiste = rstBus.RecordCount > 0
           'SE ACTUALIZA LA NOTA DE ENTRADA
           If Not rstBus.EOF Then
              stb1.SimpleText = Space(75) & "Actualizando producto a la nota de entrada: " & Adodbf.Recordset!producto
              cn.Execute "UPDATE DetalleNota SET Cantrec = " & Adodbf.Recordset!CantRecC & ",costo = " & Adodbf.Recordset!costo & " WHERE producto = '" & Adodbf.Recordset!producto & "' AND Clavenota = '" & Adodbf.Recordset!Clavenota & "'"
           Else
              stb1.SimpleText = Space(75) & "Agregando producto a la nota de entrada: " & Adodbf.Recordset!producto
              'SE DEBEN INCLUIR LOS QUE SE SOLICITAR, PORQUE SE CAPTURARON EN LA BODEGA
              CADENA = "INSERT INTO DetalleNota(ClaveNota,producto,cantsol,cantsolp,cantrec,cantrecp,costo) VALUES ('" & Adodbf.Recordset!Clavenota & "','" & Adodbf.Recordset!producto & "'," & Adodbf.Recordset!CantSolc & "," & Adodbf.Recordset!cantsolp & "," & Adodbf.Recordset!CantRecC & "," & Adodbf.Recordset!cantrecp & "," & Adodbf.Recordset!costo & ")"
             'MsgBox cadena
              cn.Execute CADENA
           End If
           'Stb1.Refresh
           cNotAnt = Adodbf.Recordset!Clavenota
           rstBus.Close
           Adodbf.Recordset.MoveNext
       Wend

   Adodbf.Recordset.MoveFirst
   MsgBox "SE GENERARON UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
   Print #1, "=========================================================================================="
   Print #1, "TOTAL DE PEDIDOS: "; nPed
   Close #1   ' Cierra el archivo de reporte
   Handle = Shell("NOTEPAD " & App.Path & "\LLEGACAR.TXT", 1)
   'MsgBox "SE ACTUALIZARON UN TOTAL DE " & CStr(nPed) & " PEDIDOS", vbInformation
   Set cnFoxPro = New ADODB.Connection
   cnFoxPro.ConnectionString = "Provider=MSDASQL.1;DSN=PITICODBF;SourceDB=" & cruta & ";SourceType=DBF;Exclusive=No;Initial Catalog= " & cruta
   cnFoxPro.Open
   cnFoxPro.Execute "UPDATE " & cArch & " SET Importado = 1 "
   'AdoPedidos.Refresh
   stb1.SimpleText = MenAnt
   stb1.Refresh
  Exit Sub
Error:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
   End If
End Sub

Private Sub importaenvios()
'On Error GoTo ERROR:
Dim cArch  As String
MenAnt = stbmensajes.SimpleText
Dim rstbusca As ADODB.Recordset
Set rstbusca = New ADODB.Recordset
frmAreaRecibo.cmdlg.FileName = ""
frmAreaRecibo.cmdlg.Filter = "Archivos Dbase (*.txt) | *.txt"
frmAreaRecibo.cmdlg.ShowOpen
cRutArc = frmAreaRecibo.cmdlg.FileName
If cRutArc = "" Or IsNull(cRutArc) Then
    MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
    Exit Sub
End If
rutac = cRutArc
For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
Next
cruta = Mid(cRutArc, 1, nPos)
cArch = Mid(cRutArc, nPos + 1)
cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.
'apertura del archivo
Open rutac For Input As #1
If EOF(1) Then
   MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
   Exit Sub
End If
SUC = Pedsuc(cArch)
If IsNull(SUC) Or Trim(SUC) = "" Then
      MsgBox "EL NOMBRE DEL ARCHIVO: " & cArch & "NO ESTA REGISTRADO EN EL SISTEMA", vbCritical, "ENVIOS"
      Exit Sub
End If
Dim Pedido As String
cveant = ""
Lbltrans.Visible = True
While Not EOF(1)
   COSTOT = 0
   'SE LEE EL ARCHIVO TXT
    Line Input #1, CAD
    nreg = nreg + 1
    'probar1.Value = nreg
    Lbltrans = nreg: Lbltrans.Refresh
    lblProd.Caption = "Registros procesados: " & Str(nreg): lblProd.Refresh
    pos1 = InStr(CAD, "|")
    traslado = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    fecha = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    tipo = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    costo = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    emisor = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    receptor = Trim(Mid(CAD, 1, pos1 - 1))
    'moy = cad
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    
    pos1 = InStr(CAD, "|")
    papeleria = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    frutas = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    Pan = Mid(CAD, 1, pos1 - 1)
    
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    merma = Mid(CAD, 1, pos1 - 1)
    'moy
     CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    auto = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    ajuste = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    producto = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    cantidad = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    cantidadp = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    costod = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    costodp = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    importe = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    iva = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    ieps = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    moy = CAD
    TASAt = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    venta = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    ventaP = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    foliotie = Mid(CAD, 1, pos1 - 1)
    CAD = Mid(CAD, pos1 + 1, Len(CAD))
    pos1 = InStr(CAD, "|")
    Pedido = CAD 'Mid(cad, 1, pos1 - 1)
    'ventap = Mid(cad, 1, pos1 - 1)
    'cad = Mid(cad, pos1 + 1, Len(cad))
    'VERIFICAR QUE NO EXISTA EL TRASLADO
    'EN EL CASO DE QUE YA EXISTA NO PASA NADA
   stbmensajes.Panels(1).Text = "Procesando traslado " & traslado & " producto " & producto
   stbmensajes.Refresh
  'If receptor = "14" Or receptor = "15" Or receptor = "27" Then
    If cveant <> traslado Then
        CADENA = "select * from traslados where t_clave = '" & Trim(traslado) & "'"
        rstbusca.Open CADENA, cn, adOpenKeyset, adLockOptimistic, adCmdText
        If rstbusca.EOF Then
            'SE INSERTA EL ENVIO
            perenv = 0: perrec = 0
            'costogral = costo
            enviado = 1: recibido = 0: entrada = 0
            devo = 0
            TRAS = 1
            CADENA = "insert into TRASLADOS (T_CLAVE,T_FECHA,T_TIPO,T_PERENV,T_PERREC,T_COSTO,T_SUCURSALEMISOR,T_SUCURSALRECEPTOR,T_ENVIADO,T_RECIBIDO, T_ENTRADA,T_FOLIOTIE,T_DEVOLUCION,T_PAPELERIA,T_FRUTAS,T_PAN,T_MERMA,T_AJUSTE,T_OBSERVA,t_perfle) VALUES( " & _
            "'" & Trim(traslado) & "','" & fecha & "'," & tipo & "," & perenv & "," & perrec & "," & costo & "," & emisor & "," & receptor & "," & enviado & "," & recibido & "," & entrada & "," & IIf(Trim(foliotie) = "", 0, foliotie) & "," & devo & "," & papeleria & "," & frutas & "," & Pan & "," & merma & "," & ajuste & ",'" & Trim(OBSERVA) & "',1)"
            'MsgBox CADENA
            cn.Execute CADENA
        End If
        rstbusca.Close
    End If
    'AHORA SE BUSCA LA PARTE DEL DETALLE POR PRODUCTO
    CADENA = "select * from detalletraslado where dt_clave = '" & Trim(traslado) & "' and dt_producto = '" & Trim(producto) & "'"
    'MsgBox cadena
    rstbusca.Open CADENA, cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rstbusca.EOF Then
       costop = 0
       'SE INSERTA EL detalle del envio
       CADENA = "insert into detalletraslado(dt_clave,dt_producto,dt_cantidad,dt_cantidadp,dt_pedido,dt_costo,dt_costop,dt_importe,dt_venta,dt_ventap,dt_iva,dt_ieps,dt_tasaieps) values(" & _
                "'" & Trim(traslado) & "','" & Trim(producto) & "'," & cantidad & "," & cantidadp & ",'" & Trim(Pedido) & "'," & costod & "," & costodp & "," & importe & "," & venta & "," & ventaP & "," & iva & "," & ieps & "," & TASAt & ")"
       cn.Execute CADENA
    Else
        'cn.Execute "UPDATE detalletraslado set dt_cantidad =" & dt_cantidadp
    End If
    rstbusca.Close
  'End If
   cveant = traslado
Wend
stbmensajes.SimpleText = MenAnt
stbmensajes.Refresh
Close #1
MsgBox "LA IMPORTACION SE REALIZO CORRECTAMENTE", vbInformation

Exit Sub
Error:
   MsgBox Err.Description
End Sub

'Actualiza los inventarios que estan dentro del array y guarda los totales en l atabla HistInv
Private Sub Inventariosimp()
Dim cArch As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
men1 = "Este proceso actualiza los inventarios enviados por las sucursales de Mayoreo..."
men2 = "Deseas Continuar...?"
RESP = MsgBox(men1 & vbCrLf & men2, vbYesNo + vbQuestion, "EXPORTACION")
If RESP = vbNo Then
   Exit Sub
End If
Dim dtas(1 To 8) As String
dtas(1) = "INVPTO"
dtas(2) = "INVCEN"
dtas(3) = "INVMIA"
dtas(4) = "INVIST"
'SE HACE UN CICLO PARA IMPORTAR LOS INVENTARIOS DE LAS TIENDAS
ANIMA
For i = 1 To 4
 If Dir("P:\BUZON\" & dtas(i) & ".TXT") <> "" Then
   Open "P:\BUZON\" & dtas(i) & ".TXT" For Input As #1
   cArch = dtas(i)
   sucu = Pedsuc(cArch)
   If Trim(sucu) = "" Then
      MsgBox "EL ARCHIVO " & cArch & " NO ESTA REGISTRADO EN EL SISTEMA, POR LO TANTO NO PUEDE CONTINUAR", vbCritical
      Exit Sub
   End If
   nreg = 0
   probar1.Min = 0
   probar1.Max = 1000
   stbmensajes.Panels(1).Text = "importando Inventario  " & cArch & " ..."
   stbmensajes.Refresh
   lblProd.Visible = True
   'SE INICIALIZA EL INVENTARIO DE LA TIENDA CORRESPONDIENTE
   CAD = "UPDATE INVENTARIO" & Trim(sucu) & " SET INCANT = 0,INCANTPZA = 0"
   cn.Execute CAD
   While Not EOF(1)
        Line Input #1, CAD
        nreg = nreg + 1
        'ProBar1.Value = nreg
        lblProd.Caption = "Registros procesados: " & Str(nreg): lblProd.Refresh
        pos1 = InStr(CAD, "|")
        clave = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        pos1 = InStr(CAD, "|")
        cajas = Mid(CAD, 1, pos1 - 1)
        CAD = Mid(CAD, pos1 + 1, Len(CAD))
        piezas = CAD
        'AHORA SE CONTRUYE LA SENTENCIA SE UTILIZA UN RECORDSET TEMPORAL
        CAD = " SELECT * FROM INVENTARIO" & Trim(sucu) & " WHERE INPROD = '" & Trim(clave) & "'"
        rs.Open CAD, cn, adOpenKeyset, adLockOptimistic, adCmdText
        If rs.EOF Then
           'SE AGREGA EL PRODUCTO
           lblProd.Caption = "Agregando ..." & clave
           lblProd.Refresh
           rs.AddNew
           rs!Inprod = clave
           rs!InCant = cajas
           rs!InCantPza = piezas
           rs!instock = 0
           rs!Insucursal = 0
           rs!InInicial = 0
           rs!ININICIALP = 0
           rs.Update
       Else
           rs!InCant = cajas
           rs!InCantPza = piezas
           rs.Update
      End If
      rs.Close
   Wend
   Close #1
  End If
Next
lblProd.Visible = False
probar1.Enabled = False
probar1.Visible = False
MsgBox "LAIMPORTACION SE REALIZO CORRECTAMENTE", vbInformation
Cmdopt_Click 1
End Sub

Private Sub cortesimp()
Dim cArch As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
men1 = "Deseas continuar con la importacion de cortes..."
RESP = MsgBox(men1, vbYesNo + vbQuestion, "EXPORTACION")
If RESP = vbNo Then
   Exit Sub
End If
'On Error GoTo ERROR:
nreg = 0
probar1.Min = 0
probar1.Max = 4000
stbmensajes.Panels(1).Text = "importando Inventario ... "
stbmensajes.Refresh
lblProd.Visible = True
'SE INICIALIZA EL INVENTARIO
Dim dtas(1 To 8) As String
dtas(1) = "CORPTO"
dtas(2) = "CORCEN"
dtas(3) = "CORMIA"
dtas(4) = "CORIST"
'SE HACE UN CICLO PARA IMPORTAR LOS INVENTARIOS DE LAS TIENDAS
ANIMA
For i = 1 To 4
  If Dir("P:\BUZON\" & dtas(i) & ".TXT") <> "" Then
     Open "P:\BUZON\" & dtas(i) & ".TXT" For Input As #1
     cArch = dtas(i)
     sucu = Pedsuc(cArch)
     stbmensajes.Panels(1).Text = "importando corte  " & cArch & " ..."
     stbmensajes.Refresh
     lblProd.Visible = True
     nreg = 1
     Line Input #1, CAD
     'DFECINI = Mid(CAD, Len(CAD) - 7, 8)
     dfecinI = IIf(InStr(1, Mid(CAD, Len(CAD) - 7, 8), "|") > 0, Mid(CAD, Len(CAD) - 10, 8), Mid(CAD, Len(CAD) - 7, 8))
     While Not EOF(1)
       Line Input #1, CAD
       nreg = nreg + 1
     Wend
     Close #1
     dfecfin = IIf(InStr(1, Mid(CAD, Len(CAD) - 7, 8), "|") > 0, Mid(CAD, Len(CAD) - 10, 8), Mid(CAD, Len(CAD) - 7, 8))
     
     probar1.Min = 0
     probar1.Max = nreg
     nreg = 0
     'SE INICIALIZA EL INVENTARIO DE LA TIENDA CORRESPONDIENTE
     If dtas(i) = "CORCEN" Then
        'CONDSER = "SERIE = 'D2' AND (CAST(factura AS MONEY) <  11100 OR (CAST(factura AS MONEY) BETWEEN  11301 AND 11500) OR (CAST(factura AS MONEY) BETWEEN  11601 AND 11699) OR CAST(factura AS MONEY) > 11901)"
        CONDSER = "SERIE = 'D2'"
     ElseIf dtas(i) = "CORPTO" Then
        CONDSER = "(SERIE = 'G2' OR SERIE = 'H2' OR SERIE = 'DDD')"
     ElseIf dtas(i) = "CORCOS" Then
        CONDSER = " (SERIE = 'I2' OR SERIE = 'J2') OR ( (SERIE = 'D2' AND factura >= 11100 AND factura <= 11300) OR (SERIE = 'D2' AND factura >= 11501 AND factura <= 11600) OR (SERIE = 'D2' AND factura >= 11700 AND factura <= 11900 ) OR (SERIE = 'D2' AND factura >= 12083 AND factura <= 12090 ) )"
     ElseIf dtas(i) = "CORMIA" Then
        'CONDSER = " ((SERIE = 'B' AND factura >= 87800 AND factura <= 87900) OR (SERIE = 'GGG')) "
        CONDSER = " (SERIE = 'B' OR SERIE = 'GGG' OR SERIE = 'HHH') "
     ElseIf dtas(i) = "CORIST" Then
        CONDSER = " (SERIE = 'D' OR SERIE = 'JJJ' OR SERIE = 'KKK' OR SERIE = 'LLL') "
     End If
     'MsgBox "DELETE FROM facvtamay WHERE fecha_det >= '" & DFECINI & "' AND fecha_det <= '" & dfecfin & "' AND " & CONDSER
     cn.Execute "DELETE FROM facvtamay WHERE fecha_det >= '" & dfecinI & "' AND fecha_det <= '" & dfecfin & "' AND " & CONDSER
     Open "P:\BUZON\" & dtas(i) & ".TXT" For Input As #1
     While Not EOF(1)
        Line Input #1, CAD
        nreg = nreg + 1
        probar1.Value = nreg
        lblProd.Caption = "Registros procesados: " & Str(nreg): lblProd.Refresh
         pos1 = InStr(CAD, "|")
        producto = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        cantidad = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        cantidadp = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        PRECIO = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        preciop = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        costo = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        costop = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        importe = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        iva = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        ieps = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        tasaieps = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        Factura = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        SERIE = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        venta = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        rfc = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        fecha = Mid(CAD, 1, 8)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
        sucu = IIf(Len(CAD) > 0, Mid(CAD, InStr(CAD, "|") + 1, Len(CAD)), "")
        'AHORA SE CONTRUYE LA SENTENCIA SE UTILIZA UN RECORDSET TEMPORAL
        cadsql = "INSERT INTO FacVtaMay VALUES ('" & producto & "'," & cantidad & "," & cantidadp & "," & PRECIO & "," & preciop & _
                   "," & costo & "," & costop & "," & importe & "," & iva & "," & ieps & "," & tasaieps & ",'" & Factura & "','" & SERIE & "'," & venta & _
                   ",'" & fecha & "','" & rfc & "'," & sucu & ")"
        cn.Execute cadsql
     Wend
     Close #1
  End If
Next
lblProd.Visible = False
probar1.Enabled = False
probar1.Visible = False
ani1.Close: ani1.Visible = False
MsgBox "LA IMPORTACION DE CORTES SE REALIZO CORRECTAMENTE", vbInformation
End Sub

Private Sub Cmdutiofi_Click(Index As Integer)
Select Case Index
Case 0
     'Importa
Case 1
      If VERFEC Then exportacambios
Case 2
     If VERFEC Then
        If MsgBox("DESEAS IMPRIMIR ETIQUETAS EN EL NUEVO FORMATO" & Chr(13) & "CON PRECIOS PIEZA Y CAJA?", vbYesNo + vbQuestion, "Formato de etiqueta") = vbYes Then
           'GenEtiqueta
           ImpEtiNva
        Else
           GENETIQUETA1
        End If
     End If
Case 3
     exportatodo
Case 4
    If MsgBox("DESEAS GENERAR INFORMACION PARA PREVENTISTAS CON EQUIPOS HP HANDHELD", vbQuestion + vbYesNo) = vbYes Then
       Dim N As Integer
       For N = 0 To lstrutas.ListCount - 1
           lstrutas.Selected(N) = True
       Next
       FRArutas.Visible = True
    Else  'Preventistas con Laptop
       cmdlg.FileName = App.Path & "\PITICO*.MDB"
       cmdlg.DialogTitle = "Búsqueda de archivo para enviar información para preventistas (Laptop)"
       cmdlg.Filter = "Archivo para preventistas (PITICO*.mdb) | PITICO*.mdb"
       cmdlg.ShowOpen
       cRutArc = cmdlg.FileName
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
       cvesuc = Trim(Mid(cSucursal, 1, 3))

       
       Dim rs As ADODB.Recordset
       Set CNMDB = New ADODB.Connection
       CNMDB.Open "DSN=PITICOMDB;DBQ=" & cruta & "\PITICO" & cvesuc & ".mdb;DefaultDir=" & App.Path & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"
       Lbltrans.Visible = True
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenStatic
       lbltransO.Visible = True
       cmen = Me.stbmensajes.Panels(1).Text

       nporcent = InputBox("Porcentaje de incremento en precios", "Incremento", 0)
       If Not IsNumeric(nporcent) Then Exit Sub
       'tabla CATCLIENTE
       stbmensajes.Panels(1).Text = "Exportando catálogo de clientes..."
       stbmensajes.Refresh
       If Trim(RutPort) = "" Then
          rs.Open "SELECT * FROM catcliente WHERE (ruta IS NULL OR RUTA = '' ) or ( CTIPO = 1 OR CTIPO = 2)", cn
       Else
          rs.Open "SELECT * FROM catcliente WHERE Substring(ruta,1,3) = '" & Trim(RutPort) & "' OR ( CTIPO = 1 OR CTIPO = 2)", cn
       End If
       Call exptablamdb(rs, "catcliente", 2)
       CNMDB.Execute "UPDATE catcliente SET cambpre = 0"
       'Tabla INVENTARIO
       stbmensajes.Panels(1).Text = "Exportando catálogo de inventario..."
       stbmensajes.Refresh
       rs.Close
       rs.Open "SELECT * FROM inventario WHERE incant > 0 or incantpza > 0", cn
       Call exptablamdb(rs, "inventario")
       'Tabla PREPROD
       stbmensajes.Panels(1).Text = "Exportando catálogo de precios..."
       stbmensajes.Refresh
       rs.Close
       rs.Open "SELECT * FROM preprod,inventario WHERE preclave = inprod And (incant > 0 or incantpza > 0)", cn
       Call exptablamdb(rs, "preprod")
       'TABLA TFPRODUC
       stbmensajes.Panels(1).Text = "Exportando catálogo de productos..."
       stbmensajes.Refresh
       rs.Close
       rs.Open "SELECT * FROM tfproduc,inventario WHERE consec = inprod And (incant > 0 or incantpza > 0)", cn
       Call exptablamdb(rs, "tfproduc")
       'TABLA FACVENTA
       stbmensajes.Panels(1).Text = "Exportando facturas..."
       stbmensajes.Refresh
       rs.Close
       rs.Open "SELECT * FROM facventa WHERE cobrado = 0 and facfecha >= '01/01/2002'", cn
       Call exptablamdb(rs, "Facventa")
       'CNMDB.Execute "DELETE FROM ventas_det"
       'CNMDB.Execute "DELETE FROM ventas"
       nporcent = 1 + (nporcent / 100)
       CNMDB.Execute "UPDATE preprod SET precio1 = precio1 * " & nporcent & ", precio2 = precio2 * " & nporcent & ", precio3 = precio3 * " & nporcent & ", precio4 = precio4 * " & nporcent & ", precio5 = precio5 * " & nporcent & ", precio6 = precio6 * " & nporcent
       stbmensajes.Panels(1).Text = cmen
       
       stbmensajes.Refresh
       MsgBox "LA EXPORTACION SE REALIZO CORRECTAMENTE", vbInformation, "Exportación"
    End If
End Select
End Sub

Private Sub ImpEtiNva()
Dim rs As ADODB.Recordset
On Error GoTo Error:
Dim cImpAct As String
Set rs = New ADODB.Recordset
txtinicio1 = DateAdd("d", -1, txtinicio.Text)
txtfinal1 = DateAdd("d", 1, txtfinal.Text)
cImpAct = ""
cresp = MsgBox("DESEAS IMPRIMIR LAS ETIQUETAS DE PRODUCTOS INACTIVOS", vbYesNoCancel + vbQuestion + vbDefaultButton2)
If cresp = vbNo Then
   cImpAct = " AND activo = 1 "
ElseIf cresp = vbCancel Then
   Exit Sub
End If
rs.Open "SELECT flesub, fletex,CONSEC,descripc,paquetes, RTRIM(str(contenid,10,3)) + ' ' + rtrim(mediDa)  AS MEDIDA , precio1,precio2,precio3 FROM TFPRODUC,PREPROD WHERE consec = preclave " & cImpAct & compInt & _
    "AND fecact > '" & txtinicio1 & " ' and fecact < '" & txtfinal1 & " ' ORDER BY DESCRIPC ", cn, adOpenKeyset, adLockOptimistic, adCmdText
'rs.Open "SELECT descripc, RTRIM(str(contenid,10,3)) + ' ' + mediDa  AS MEDIDA , precio1 FROM TFPRODUC,PREPROD,EnvioMina WHERE consec = preclave " & cImpAct & _
'    "AND consec = producto AND preclave = producto ORDER BY DESCRIPC ", cn, adOpenKeyset, adLockOptimistic, adCmdText

MsgBox "A CONTINUACION SE IMPRIMIRAN ETIQUETAS DE " & rs.RecordCount & " PRODUCTOS", vbInformation
cmdlg.DialogTitle = "Páginas a Imprimirse "
cmdlg.CancelError = True
cmdlg.Flags = &H100000   'Oculta la casilla de imprimir en archivoo
cmdlg.ShowPrinter
nTotEti = 0
For x = 1 To cmdlg.Copies
 rs.MoveFirst
 Printer.ScaleMode = vbCentimeters
 Printer.CurrentX = 1
 'Printer.Font = "ARIAL BLACK": Printer.FontSize = "12"
 Printer.Font = "ARIAL NARROW": Printer.FontSize = "12"
 nposx = 0: nPosy = 0: nProd = 0: N = 0
 Printer.Line (0, nPosy)-(21.5, nPosy)
 cMens = stbmensajes.Panels(1).Text
 While Not rs.EOF
    Printer.FontSize = "12"
    descr = Trim(rs!descripc)
    If Printer.TextWidth(descr) > 9.5 Then
       Printer.CurrentY = 0.1 + nPosy
       nPos = Int(Len(descr) / 2)
       nComp = 0
       Do  'Se imprimen palabras completas
         nComp = InStr(nComp + 1, descr, " ")
         If nComp > 0 Then espacio = nComp
       Loop Until nComp >= nPos Or nComp = 0
       nPos = espacio - 1
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(Mid(descr, 1, nPos))) / 2  'Centro el dato
       Printer.Print Mid(descr, 1, nPos)
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(Trim(Mid(descr, nPos + 1)) + "  " + rs!medida)) / 2 'Centro el dato
       Printer.Print Trim(Mid(descr, nPos + 1) + "  " + rs!medida)
    Else  'Si cabe en una sola linea
       Printer.CurrentY = 0.3 + nPosy
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(descr)) / 2
       Printer.Print Trim(rs!descripc)
       Printer.CurrentX = nposx + 3.3
       Printer.Print rs!medida
    End If
    Printer.CurrentX = nposx + 1
    Printer.CurrentY = 1.2 + nPosy
    Printer.Print "Cant." & Space(26) & "Precio Unit." & Space(12) & "Importe"
    Printer.FontSize = "16"
    If rs!PAQUETES = 1 Then Printer.CurrentY = 2.2 + nPosy
    'Precio por pieza
    Printer.CurrentX = nposx + 0.5
    Printer.Print "(1) Pieza";
    Printer.CurrentX = nposx + 4.2
    Printer.Print Format(rs!precio1, "$ ####,##0.00");
    Printer.CurrentX = nposx + 6.8
    Printer.Print Format(rs!precio1, "$ ####,##0.00")
    If rs!PAQUETES > 1 Then
          Printer.CurrentY = 2.5 + nPosy
          'Precio por caja
          Printer.CurrentX = nposx + 0.5
          Printer.Print Format(rs!PAQUETES, "(###)"); " Pzas (caja)";
          Printer.CurrentX = nposx + 4.2
          Printer.Print Format(rs!PRECIO2 / rs!PAQUETES, "$ ####,##0.00");
          Printer.CurrentX = nposx + 6.8
          Printer.Print Format(rs!PRECIO2, "$ ####,##0.00")
    End If
    nposx = 10.5
    'Printer.Font = "ARIAL BLACK": Printer.FontSize = "12"
    If (N + 1) Mod 2 = 0 Then
       nPosy = nPosy + 3.8
       Printer.DrawWidth = 15
       Printer.Line (10.3, nPosy)-(20.5, nPosy)
       Printer.Line (10.3, nPosy - 3.5)-(10.3, nPosy)
       nposx = 0
    Else
       Printer.DrawWidth = 15
       Printer.Line (0, nPosy + 3.8)-(10, nPosy + 3.8)
       Printer.DrawWidth = 30
       Printer.Line (0, nPosy + 3.8 - 3.5)-(0, nPosy + 3.8)
    End If
    nProd = nProd + 1
    nTotEti = nTotEti + 1
    stbmensajes.Panels(1).Text = "Imprimiendo etiqueta " & nTotEti & " de " & rs.RecordCount
    stbmensajes.Refresh
    If nProd = 14 Then
       nProd = 0
       Printer.EndDoc
       Printer.ScaleMode = vbCentimeters
       Printer.CurrentX = 1
       Printer.Font = "ARIAL NARROW": Printer.FontSize = "12"
       nposx = 0: nPosy = 0: nProd = 0
    End If
    N = N + 1
    rs.MoveNext
 Wend
 'Printer.Line (0, nPosy)-(21, nPosy)
 Printer.EndDoc
Next
stbmensajes.Panels(1).Text = cMens
stbmensajes.Refresh
Exit Sub
Error:
  If Err.Number <> 32755 Then
     MsgBox Err.Description
  End If
End Sub

Private Sub GenEtiqueta()
Dim rs As ADODB.Recordset
On Error GoTo Error:
Dim cImpAct As String
Set rs = New ADODB.Recordset
txtinicio1 = DateAdd("d", -1, txtinicio.Text)
txtfinal1 = DateAdd("d", 1, txtfinal.Text)
cImpAct = ""
cresp = MsgBox("DESEAS IMPRIMIR LAS ETIQUETAS DE PRODUCTOS INACTIVOS", vbYesNoCancel + vbQuestion + vbDefaultButton2)
If cresp = vbNo Then
   cImpAct = " AND activo = 1 "
ElseIf cresp = vbCancel Then
   Exit Sub
End If
rs.Open "SELECT flesub, fletex,CONSEC,descripc,paquetes, RTRIM(str(contenid,10,3)) + ' ' + rtrim(mediDa)  AS MEDIDA , precio1,precio2,precio3 FROM TFPRODUC,PREPROD WHERE consec = preclave " & cImpAct & compInt & _
    "AND fecact > '" & txtinicio1 & " ' and fecact < '" & txtfinal1 & " ' ORDER BY DESCRIPC ", cn, adOpenKeyset, adLockOptimistic, adCmdText
'rs.Open "SELECT descripc, RTRIM(str(contenid,10,3)) + ' ' + mediDa  AS MEDIDA , precio1 FROM TFPRODUC,PREPROD,EnvioMina WHERE consec = preclave " & cImpAct & _
'    "AND consec = producto AND preclave = producto ORDER BY DESCRIPC ", cn, adOpenKeyset, adLockOptimistic, adCmdText

MsgBox "A CONTINUACION SE IMPRIMIRAN ETIQUETAS DE " & rs.RecordCount & " PRODUCTOS", vbInformation
cmdlg.DialogTitle = "Páginas a Imprimirse "
cmdlg.CancelError = True
cmdlg.Flags = &H100000   'Oculta la casilla de imprimir en archivoo
cmdlg.ShowPrinter
nTotEti = 0
For x = 1 To cmdlg.Copies
 rs.MoveFirst
 Printer.ScaleMode = vbCentimeters
 Printer.CurrentX = 1
 'Printer.Font = "ARIAL BLACK": Printer.FontSize = "12"
 Printer.Font = "ARIAL NARROW": Printer.FontSize = "12"
 nposx = 0: nPosy = 0: nProd = 0: N = 0
 Printer.Line (0, nPosy)-(21.5, nPosy)
 cMens = stbmensajes.Panels(1).Text
 While Not rs.EOF
    Printer.FontSize = "12"
    descr = Trim(rs!descripc)
    If Printer.TextWidth(descr) > 9.5 Then
       Printer.CurrentY = 0.1 + nPosy
       nPos = Int(Len(descr) / 2)
       nComp = 0
       Do  'Se imprimen palabras completas
         nComp = InStr(nComp + 1, descr, " ")
         If nComp > 0 Then espacio = nComp
       Loop Until nComp >= nPos Or nComp = 0
       nPos = espacio - 1
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(Mid(descr, 1, nPos))) / 2  'Centro el dato
       Printer.Print Mid(descr, 1, nPos)
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(Trim(Mid(descr, nPos + 1)) + "  " + rs!medida)) / 2 'Centro el dato
       Printer.Print Trim(Mid(descr, nPos + 1) + "  " + rs!medida)
    Else  'Si cabe en una sola linea
       Printer.CurrentY = 0.3 + nPosy
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(descr)) / 2
       Printer.Print Trim(rs!descripc)
       Printer.CurrentX = nposx + 3.3
       Printer.Print rs!medida
    End If
    Printer.CurrentX = nposx + 1
    Printer.CurrentY = 1.2 + nPosy
    Printer.Print "Cant." & Space(26) & "Precio Unit." & Space(12) & "Importe"
    Printer.FontSize = "16"
    If rs!PAQUETES = 1 Then Printer.CurrentY = 2.2 + nPosy
    'Precio por pieza
    Printer.CurrentX = nposx + 1
    Printer.Print "(1) Pieza";
    Printer.CurrentX = nposx + 4.2
    Printer.Print Format(rs!precio1, "$ ####,##0.00");
    Printer.CurrentX = nposx + 6.8
    Printer.Print Format(rs!precio1, "$ ####,##0.00")
    If rs!PAQUETES > 1 Then
       If rs!PAQUETES >= 10 Then
          'Precio por Medio Mayoreo
           pzamm = rs!flesub
           Printer.CurrentX = nposx + 1
           'Printer.Print Format(pzamm, "(###)"); " Pack";
           Printer.Print Format(pzamm, "(###)"); " Pzas.";
           Printer.CurrentX = nposx + 4.2
           Printer.Print Format(rs!fletex, "$ ####,##0.00");
           Printer.CurrentX = nposx + 6.8
           Printer.Print Format(pzamm * rs!fletex, "$ ####,##0.00")
       Else
          Printer.CurrentY = 2.5 + nPosy
       End If
          'Precio por caja
          Printer.CurrentX = nposx + 1
          Printer.Print Format(rs!PAQUETES, "(###)"); " Pzas.";
          Printer.CurrentX = nposx + 4.2
          Printer.Print Format(rs!PRECIO2 / rs!PAQUETES, "$ ####,##0.00");
          Printer.CurrentX = nposx + 6.8
          Printer.Print Format(rs!PRECIO2, "$ ####,##0.00")
    End If
    nposx = 10.5
    'Printer.Font = "ARIAL BLACK": Printer.FontSize = "12"
    If (N + 1) Mod 2 = 0 Then
       nPosy = nPosy + 3.8
       Printer.DrawWidth = 15
       Printer.Line (10.3, nPosy)-(20.5, nPosy)
       Printer.Line (10.3, nPosy - 3.5)-(10.3, nPosy)
       nposx = 0
    Else
       Printer.DrawWidth = 15
       Printer.Line (0, nPosy + 3.8)-(10, nPosy + 3.8)
       Printer.DrawWidth = 30
       Printer.Line (0, nPosy + 3.8 - 3.5)-(0, nPosy + 3.8)
    End If
    nProd = nProd + 1
    nTotEti = nTotEti + 1
    stbmensajes.Panels(1).Text = "Imprimiendo etiqueta " & nTotEti & " de " & rs.RecordCount
    stbmensajes.Refresh
    If nProd = 14 Then
       nProd = 0
       Printer.EndDoc
       Printer.ScaleMode = vbCentimeters
       Printer.CurrentX = 1
       Printer.Font = "ARIAL NARROW": Printer.FontSize = "12"
       nposx = 0: nPosy = 0: nProd = 0
    End If
    N = N + 1
    rs.MoveNext
 Wend
 'Printer.Line (0, nPosy)-(21, nPosy)
 Printer.EndDoc
Next
stbmensajes.Panels(1).Text = cMens
stbmensajes.Refresh
Exit Sub
Error:
  If Err.Number <> 32755 Then
     MsgBox Err.Description
  End If
End Sub

Private Sub GENETIQUETA1()
Dim rs As ADODB.Recordset
On Error GoTo Error:
Dim cImpAct As String
Set rs = New ADODB.Recordset
txtinicio1 = DateAdd("d", -1, txtinicio.Text)
txtfinal1 = DateAdd("d", 1, txtfinal.Text)
cImpAct = ""
cresp = MsgBox("DESEAS IMPRIMIR LAS ETIQUETAS DE PRODUCTOS INACTIVOS", vbYesNoCancel + vbQuestion + vbDefaultButton2)
If cresp = vbNo Then
   cImpAct = " AND activo = 1 "
ElseIf cresp = vbCancel Then
   Exit Sub
End If
rs.Open "SELECT CONSEC,descripc,paquetes, RTRIM(str(contenid,10,3)) + ' ' + rtrim(mediDa)  AS MEDIDA , precio1,precio2,precio3 FROM TFPRODUC,PREPROD WHERE consec = preclave " & cImpAct & compInt & _
    " AND fecact > '" & txtinicio1 & " ' and fecact < '" & txtfinal1 & " ' ORDER BY DESCRIPC", cn, adOpenKeyset, adLockOptimistic, adCmdText
'rs.Open "SELECT descripc, RTRIM(str(contenid,10,3)) + ' ' + mediDa  AS MEDIDA , precio1 FROM TFPRODUC,PREPROD,BREOCTNOV WHERE rtrim(consec) = rtrim(preclave) " & _
    "AND rtrim(consec) = rtrim(producto) AND rtrim(preclave) = rtrim(producto) ORDER BY DESCRIPC ", cn, adOpenKeyset, adLockReadOnly, adCmdText
MsgBox "A CONTINUACION SE IMPRIMIRAN ETIQUETAS DE " & rs.RecordCount & " PRODUCTOS", vbInformation
cmdlg.DialogTitle = "Páginas a Imprimirse "
cmdlg.CancelError = True
cmdlg.Flags = &H100000   'Oculta la casilla de imprimir en archivoo
cmdlg.ShowPrinter
nTotEti = 0
For x = 1 To cmdlg.Copies
 rs.MoveFirst
 Printer.ScaleMode = vbCentimeters
 Printer.CurrentX = 1
 Printer.Font = "ARIAL BLACK": Printer.FontSize = "12"
 nposx = 0: nPosy = 0: nProd = 0: N = 0
 Printer.Line (0, nPosy)-(21.5, nPosy)
 cMens = stbmensajes.Panels(1).Text
 While Not rs.EOF
    descr = Trim(rs!descripc)
    If Printer.TextWidth(descr) > 9.5 Then
       Printer.CurrentY = 0.1 + nPosy
       nPos = Int(Len(descr) / 2)
       nComp = 0
       Do  'Se imprimen palabras completas
         nComp = InStr(nComp + 1, descr, " ")
         If nComp > 0 Then espacio = nComp
       Loop Until nComp >= nPos Or nComp = 0
       nPos = espacio - 1
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(Mid(descr, 1, nPos))) / 2  'Centro el dato
       Printer.Print Mid(descr, 1, nPos)
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(Trim(Mid(descr, nPos + 1)))) / 2 'Centro el dato
       Printer.Print Trim(Mid(descr, nPos + 1))
    Else  'Si cabe en una sola linea
       Printer.CurrentY = 0.3 + nPosy
       Printer.CurrentX = nposx + (10 - Printer.TextWidth(descr)) / 2
       Printer.Print Trim(rs!descripc)
    End If
    Printer.CurrentX = nposx + 3.3
    Printer.Print rs!medida
    Printer.Font = "haettenschweiler": Printer.FontSize = "50"
    Printer.CurrentX = nposx + 3
    Printer.Print Format(rs!precio1, "$ ####,##0.00")
    nposx = 10.5
    Printer.Font = "ARIAL BLACK": Printer.FontSize = "12"
    If (N + 1) Mod 2 = 0 Then
       nPosy = nPosy + 3.8
       Printer.DrawWidth = 15
       Printer.Line (10.3, nPosy)-(20.5, nPosy)
       Printer.Line (10.3, nPosy - 3.5)-(10.3, nPosy)
       nposx = 0
    Else
       Printer.DrawWidth = 15
       Printer.Line (0, nPosy + 3.8)-(10, nPosy + 3.8)
       Printer.DrawWidth = 30
       Printer.Line (0, nPosy + 3.8 - 3.5)-(0, nPosy + 3.8)
    End If
    nProd = nProd + 1
    nTotEti = nTotEti + 1
    stbmensajes.Panels(1).Text = "Imprimiendo etiqueta " & nTotEti & " de " & rs.RecordCount
    stbmensajes.Refresh
    If nProd = 14 Then
       nProd = 0
       Printer.EndDoc
       Printer.ScaleMode = vbCentimeters
       Printer.CurrentX = 1
       Printer.Font = "ARIAL BLACK": Printer.FontSize = "12"
       nposx = 0: nPosy = 0: nProd = 0
    End If
    N = N + 1
    rs.MoveNext
 Wend
 'Printer.Line (0, nPosy)-(21, nPosy)
 Printer.EndDoc
Next
stbmensajes.Panels(1).Text = cMens
stbmensajes.Refresh
Exit Sub
Error:
  If Err.Number <> 32755 Then
     MsgBox Err.Description
  End If
End Sub


'Función que verifica que el rango de fechas sea correcto
Function VERFEC() As Boolean
On Error GoTo Error:
Dim fecini As Date
Dim fecfin As Date
If Not fraperi.Visible Then
   fraperi.Visible = True
   txtinicio.SetFocus
   VERFEC = False
Else
   fecini = txtinicio.Text
   fecfin = txtfinal.Text
   If fecini > fecfin Then
      MsgBox "Rango de Fechas Invalidas..."
      txtinicio.SetFocus
      VERFEC = False
   Else
      VERFEC = True
   End If
End If
Error:
    If Err.Description <> "" Then MsgBox Err.Description
End Function

Private Sub CmdutiTer_Click()
    fraUtiOfi.Visible = False
    fraUtiOfi.Refresh
    'Me.cmdBotones(7).SetFocus
End Sub

Private Sub cmdverofertas_Click()
CR1.WindowTitle = "Listado de productos ofertados"
CR1.ReportFileName = App.Path & "\ofertas.rpt"
CR1.DataFiles(0) = "preciostemp"
CR1.Formulas(0) = ""
CR1.Connect = strconnect
If MsgBox("DESEAS GENERAR ARCHIVO DE OFERTAS PARA TIENDAS", vbYesNo + vbQuestion + vbDefaultButton2, "Compras") = vbYes Then
    CR1.Destination = crptToFile
    CR1.PrintFileName = "p:\paso\ofertas.txt"
    CR1.PrintFileType = crptText
    MsgBox "El Archivo Ofertas.txt se encuentra en P:\Paso", vbInformation, "Ruta"
Else
    CR1.Destination = crptToWindow
End If
CR1.Action = 1
Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub Command1_Click()
' Call precycosto
End Sub

Private Sub precycosto()
Dim rs  As ADODB.Recordset
Dim rsall As ADODB.Recordset
Set rsall = New ADODB.Recordset
stbmensajes.SimpleText = "Espere Generando Informacion..."
stbmensajes.Refresh
Dim tempo As Double
Open "c:\paso\difcosto.txt" For Output As #1
Open "c:\paso\difpre.txt" For Output As #2
cn.Execute "delete malos"
CAD = "select * from tfproduc WHERE activo = 1 AND ofertado = 0 AND interno = 0 AND medmay = 0"
rsall.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
Set rs = New ADODB.Recordset
N = 0
While Not rsall.EOF
   N = N + 1
   If Trim(rsall!CONSEC) = "1003895" Then MsgBox "asad"

    cajas = rsall!cajas
    encajas = rsall!encajas
    costo = rsall!PRECOSTO
    PAQUETES = rsall!PAQUETES
    tempo = 0
    tt = 0
    CONSEC = rsall!CONSEC
    clave = rsall!CONSEC
    PRELISTA = rsall!costocaj
    tempo = PRELISTA
    flete = 0
    otrocargo = 0
    CAD = "select * from cargos  where caprod =  '" & Trim(clave) & "'"
    rs.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Not rs.EOF Then
        tt = (PRELISTA * (IIf(Not IsNull(rs!cargo1), rs!cargo1, 0) / 100))
        tempo = PRELISTA + tt
        tt = (tempo * (IIf(Not IsNull(rs!cargo2), rs!cargo2, 0) / 100))
        tempo = tempo + tt
        tempo = tempo + (tempo * (IIf(Not IsNull(rs!ieps), rs!ieps, 0) / 100))
        tempo = tempo + (tempo * (IIf(Not IsNull(rs!iva), rs!iva, 0) / 100))
        
        tempo = tempo + IIf(Not IsNull(rs!cargo_efectivo), rs!cargo_efectivo, 0)
        
        flete = IIf(Not IsNull(rs!flete_efectivo), rs!flete_efectivo, 0)
        otrocargo = IIf(Not IsNull(rs!maniobras), rs!maniobras, 0)
    End If
    rs.Close
    CAD = "select * from descuentos where deprod = '" & Trim(clave) & "'"
    rs.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        tempo = tempo - (tempo * (IIf(Not IsNull(rs!decto1), rs!decto1, 0) / 100))
        tempo = tempo - (tempo * (IIf(Not IsNull(rs!decto2), rs!decto2, 0) / 100))
        tempo = tempo - (tempo * (IIf(Not IsNull(rs!decto3), rs!decto3, 0) / 100))
        tempo = tempo - (tempo * (IIf(Not IsNull(rs!dectoOferta), rs!dectoOferta, 0) / 100))
        
        tempo = tempo - (tempo * (IIf(Not IsNull(rs!decto5), rs!decto5, 0) / 100))
        
        tempo = tempo - IIf(Not IsNull(rs!dectoefectivo), rs!dectoefectivo, 0)
        tempo = tempo - (tempo * (IIf(Not IsNull(rs!dectoFinanciero), rs!dectoFinanciero, 0) / 100))
        'tempo = tempo + flete
        'tempo = tempo + otrocargo
        costosinprom = tempo
    End If
     tempo = tempo + flete
     tempo = tempo + otrocargo
    If cajas > 0 And encajas > 0 Then
        Promocion = (tempo * encajas) / (encajas + cajas)
    Else
        Promocion = tempo
    End If
    
    If Abs(Promocion - costo) >= 1 Then
        CAD = "INSERT INTO MALOS(CONSEC,PRECIOCAL,PRECIOBAS,TIPO) VALUES ( '" & Trim(clave) & "'," & Promocion & "," & costo & ",'COSTO')"
        'MsgBox CAD
        cn.Execute CAD
    End If

    'se sacan los precios con base a las escalas que tiene el producto
    rs.Close
    CAD = "select * from margen  where producto = '" & Trim(clave) & "'"
    rs.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        escala1 = rs!escala1
        escala2 = rs!escala2
        escala3 = rs!escala3
        escala4 = rs!escala4
    End If
    precio1 = Round(Promocion * (1 + (escala1 / 100)) / PAQUETES)
    PRECIO2 = Round(Promocion * (1 + (escala2 / 100)))
    PRECIO3 = Round(Promocion * (1 + (escala3 / 100)))
    precio4 = Round(Promocion * (1 + (escala4 / 100)))
    rs.Close
    CAD = "select * from preprod  where preclave = '" & Trim(clave) & "'"
    rs.Open CAD, cn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        p1 = rs!precio1
        p2 = rs!PRECIO2
        p3 = rs!PRECIO3
        p4 = rs!precio4
    End If
    rs.Close
    prod = rsall!descripc & "  " & rsall!PAQUETES & " x " & rsall!CONTENID & "  " & rsall!medida
    If Abs(p1 - precio1) >= 1 Then
        CAD = "INSERT INTO MALOS(CONSEC,PRECIOCAL,PRECIOBAS,TIPO) VALUES ( '" & Trim(CONSEC) & "'," & precio1 & "," & p1 & ",'PRECIO')"
        'MsgBox CAD
        cn.Execute CAD
    End If
    Command1.Caption = CONSEC
    Command1.Refresh
    rsall.MoveNext
    'MsgBox rsall.RecordCount
    Me.stbmensajes.SimpleText = " consec " & CONSEC
    stbmensajes.Refresh
    Me.Caption = " consec " & CONSEC & "    Prod : " & CStr(N)
Wend
Close #1
Close #2
MsgBox "Ya termine"
End Sub

Private Sub Form_Activate()
On Error Resume Next
 cmdBotones(4).SetFocus
 If txtbusca.Visible Then gridbusca.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If frabusqueda.Enabled Then
        frabusqueda.Enabled = False
        frabusqueda.Visible = False
        StrTeclaPres = ""
        If cmdBotones(4).Enabled Then cmdBotones(4).SetFocus
     Else
        StrTeclaPres = ""
    End If
End If

If StrTeclaPres = "" Then
Select Case KeyCode
    Case 113
        'BUSQUEDA DE PRODUCTOS Y ENTRADA A DETALLE DE PRODUCTOS
         StrTeclaPres = "F2"
    Case 114
        'BUSQUEDA DE PRODUCTOS Y ENTRADA A PRECIOS
         StrTeclaPres = "F3"
    Case 115
        'BUSQUEDA POR CODIGO DE BARRAS
         StrTeclaPres = "F4"
    Case 116
        'BUSQUEDA POR CLAVE CONSEC
         StrTeclaPres = "F5"
    Case 117
        'BUSQUEDA POR CLAVE CONSEC
         StrTeclaPres = "F6"
    Case 119
         frmCalc.Show
    Case 120
         Fraimporta.Visible = True
         Fraimporta.Enabled = True
    Case 121
         Dim CAD As String
         Open App.Path & "\CAMBCOD.TXT" For Input As #1
         On Error GoTo Error:
         lTrans = True
         cn.BeginTrans
         While Not EOF(1)
            Line Input #1, CAD
            pos1 = InStr(CAD, ",")
            cvecab = Mid(CAD, 1, pos1 - 1)
            CAD = Mid(CAD, pos1 + 1, Len(CAD))
            cveofi = Mid(CAD, 1)
            cn.Execute "UPDATE tfproduc SET igualofi = 1, consec = '" & cveofi & "' WHERE consec = '" & cvecab & "'"
            cn.Execute "UPDATE inventario SET inprod = '" & cveofi & "' WHERE inprod = '" & cvecab & "'"
            cn.Execute "UPDATE ventas_det SET cl_producto = '" & cveofi & "' WHERE cl_producto = '" & cvecab & "'"
            cn.Execute "UPDATE facventa_det SET producto = '" & cveofi & "' WHERE producto = '" & cvecab & "'"
            cn.Execute "UPDATE detallefactura SET df_prod = '" & cveofi & "' WHERE df_prod = '" & cvecab & "'"
            cn.Execute "UPDATE detalleglobal SET dg_producto = '" & cveofi & "' WHERE dg_producto = '" & cvecab & "'"
            cn.Execute "UPDATE detallefactura SET df_prod = '" & cveofi & "' WHERE df_prod = '" & cvecab & "'"
            cn.Execute "UPDATE detallenota SET producto = '" & cveofi & "' WHERE producto = '" & cvecab & "'"
            cn.Execute "UPDATE detalletraslado SET dt_producto = '" & cveofi & "' WHERE dt_producto = '" & cvecab & "'"
            cn.Execute "UPDATE detalleback SET producto = '" & cveofi & "' WHERE producto = '" & cvecab & "'"
            cn.Execute "UPDATE detalleajustes SET da_producto = '" & cveofi & "' WHERE da_producto = '" & cvecab & "'"
            cn.Execute "UPDATE cambpre SET producto = '" & cveofi & "' WHERE producto = '" & cvecab & "'"
            cn.Execute "UPDATE cargos SET caprod = '" & cveofi & "' WHERE caprod = '" & cvecab & "'"
            cn.Execute "UPDATE descprod SET producto = '" & cveofi & "' WHERE producto = '" & cvecab & "'"
            cn.Execute "UPDATE descuentos SET deprod = '" & cveofi & "' WHERE deprod = '" & cvecab & "'"
            cn.Execute "UPDATE cargos SET caprod = '" & cveofi & "' WHERE caprod = '" & cvecab & "'"
            cn.Execute "UPDATE margen SET producto = '" & cveofi & "' WHERE producto = '" & cvecab & "'"
            cn.Execute "UPDATE preprod SET preclave = '" & cveofi & "' WHERE preclave = '" & cvecab & "'"
            cn.Execute "UPDATE preciostemp SET producto = '" & cveofi & "' WHERE producto = '" & cvecab & "'"
            cn.Execute "UPDATE preciohis SET producto = '" & cveofi & "' WHERE producto = '" & cvecab & "'"
            nreg = nreg + 1
         Wend
         Close #1
         cn.CommitTrans
         MsgBox "LA ACTUALIZACION SE REALIZO CORRECTAMENTE", vbInformation, "Cambio de códigos"
    Case Else
        StrTeclaPres = ""
End Select
If StrTeclaPres <> "" Then
    frabusqueda.Enabled = True
    frabusqueda.Visible = True
    txtbusca.Text = ""
    txtbusca.SetFocus
End If
Exit Sub
Error:
     If lTrans Then cn.CommitTrans
End If
End Sub

Private Sub CmdBusca_Click()
Me.frabusqueda.Visible = False
StrTeclaPres = ""
txtbusca.Text = ""
cmdBotones(4).SetFocus
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
'Dim mes(11) As Variant
On Error Resume Next
  mes(0) = "Enero": mes(1) = "Febrero": mes(2) = "Marzo": mes(3) = "Abril":  mes(4) = "Mayo": mes(5) = "Junio": mes(6) = "Julio"
  mes(7) = "Agosto": mes(8) = "Septiembre":  mes(9) = "Octubre": mes(10) = "Noviembre": mes(11) = "Diciembre"
  stbmensajes.Panels(1).Text = "Alt + tecla resaltada activa Menú"
  stbmensajes.Panels(2).Text = Str(Day(date)) + " de " + mes(Month(date) - 1) & " del " & Str(Year(date))
  stbmensajes.Panels(3).Text = Format(Time, "hh:mm AM/PM")
  If Not Sql Then
    For N = 0 To 10
      Me.cmdBotones(N).Enabled = False
    Next
    Me.mnupedidos.Enabled = False
    Me.MnuTraslados.Enabled = False
    Me.mnuutil.Enabled = False
    Me.mnuventas.Enabled = False
    Me.mnuinventario.Enabled = False
    cmdBotones(4).Enabled = True
    cmdBotones(9).Enabled = True
    cmdBotones(2).Enabled = True
  End If
  Set rs = New ADODB.Recordset
  rs.Open "SELECT tapiz,soloact,LEVEL1,permisos FROM USUARIOS WHERE clave = '" & Trim(Mid(cCveDesUsu, 1, 3)) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
  If Trim(rs!tapiz) <> "" Then Set Imgtapiz.Picture = LoadPicture(rs!tapiz)
  MnuSolAct.Checked = rs!SoloAct
  SoloAct = rs!SoloAct   'Solo activos
  consulta = ""
  Me.Caption = ccaption
  Unload frmLogin
  lbus = False
  flash.Playing = True
  AsignaPer rs!LEVEL1, rs!permisos 'Asigna permisos
End Sub

Private Sub AsignaPer(Nivel As String, permisos As String)
If Nivel = "V" Then   'Punto de venta
  mnupedidos.Enabled = False
  MnuTraslados.Enabled = False
  mnuutil.Enabled = False
  mnuorg.Enabled = False
  For N = 0 To 12
    If N <> 4 And N <> 9 And N <> 11 And N <> 6 Then cmdBotones(N).Enabled = False
  Next
Else
   'Traslados
   cmdBotones(1).Enabled = IIf(Mid(permisos, 4, 1) = 1, True, False)
   MnuTraslados.Enabled = IIf(Mid(permisos, 4, 1) = 1, True, False)
   'Pedidos por proveedor, abastecimiento, sugeridos
   cmdBotones(7).Enabled = IIf(Mid(permisos, 5, 1) = 1, True, False)
   mnupedidos.Enabled = IIf(Mid(permisos, 5, 1) = 1, True, False)
   mnupedaba.Enabled = IIf(Mid(permisos, 5, 1) = 1, True, False)
   cmdBotones(0).Enabled = IIf(Mid(permisos, 5, 1) = 1, True, False)
   cmdBotones(5).Enabled = IIf(Mid(permisos, 5, 1) = 1, True, False)
   'Facturas
   cmdBotones(10).Enabled = IIf(Mid(permisos, 6, 1) = 1, True, False)
   'Ventas
   cmdBotones(9).Enabled = IIf(Mid(permisos, 7, 1) = 1, True, False)
   mnuventas.Enabled = IIf(Mid(permisos, 7, 1) = 1, True, False)
   If Nivel = "B" Or Nivel = "C" Then
      cmdBotones(2).Enabled = (Nivel = "C")
     ' cmdBotones(2).Enabled = False
      cmdBotones(3).Enabled = False
      'cmdBotones(6).Enabled = False
      mnuutil.Enabled = False
      cmdBotones(8).Enabled = False
      mnuorg.Enabled = False
   End If
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
'ImgTapiz.Width = frmAreaRecibo.ScaleWidth - 200
'ImgTapiz.Height = frmAreaRecibo.ScaleHeight - 2400
FraMenu.Width = frmAreaRecibo.ScaleWidth - 200
flash.Width = frmAreaRecibo.ScaleWidth - 200
'lblEnca(0).Width = frmAreaRecibo.ScaleWidth
'lblEnca(1).Width = frmAreaRecibo.ScaleWidth
End Sub

Private Sub mnuInvMod_Click()
   frmModInv.Show
End Sub

Private Sub MnuPedCaptura_Click()
  cModo = "CAPTURARPEDIDO"
  frmCaptPed.Caption = "Captura de pedido"
  frmCaptPed.Show
End Sub

Private Sub mnuPedCapAlt_Click()
  nOp = 1
  cModo = "CAPTURARPEDIDO"
  frmCaptPed.Caption = "Capturar nuevos pedidos"
  frmCaptPed.Show
End Sub

Private Sub MnuPedcapMod_Click()
  nOp = 0
  cModo = "CAPTURARPEDIDO"
  frmCaptPed.Caption = "Modificaciones de pedidos"
  Load frmCaptPed
  frmCaptPed.Show
End Sub

Private Sub MnuPedCon_Click()
  nOp = 0
  cModo = "CONFIRMARPEDIDO"
  frmCaptPed.Caption = "Confirmar pedido"
  frmCaptPed.Show
End Sub

Private Sub MnuPedRec_Click()
   nOp = 0
   cModo = "RECIBIRPEDIDO"
   frmCaptPed.Caption = "Recibir pedido"
   frmCaptPed.Show
End Sub

Private Sub gridbusca_DblClick()
If Not Sql Then
   Importa (Trim(adobus.Recordset!CONSEC))
End If
End Sub

Private Sub gridbusca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error GoTo Error:
reg = gridbusca.Bookmark
    Select Case StrTeclaPres
        Case "F2"
            lpprod = True
            reg = gridbusca.Bookmark
            strcveprod = gridbusca.Columns(0).Text
            frmprod.Show 1
        Case "F3"
            strcveprov = adobus.Recordset!claprove
            strcveprod = gridbusca.Columns(0).Text
            lpprov = True
            lpprod = False
            If InStr(1, UCase(computadora()), "MODVTA") = 0 Or InStr(1, UCase(Mid(computadora(), 1, 6)), "PREVTA") = 0 Then
                frmprecios.Show
            Else
                fnewprec.Show
            End If
        Case "F4"
            lpprod = True
            reg = gridbusca.Bookmark
            strcveprod = gridbusca.Columns(0).Text
            frmprod.Show 1
        Case "F5"
            lpprod = True
            reg = gridbusca.Bookmark
            strcveprod = gridbusca.Columns(0).Text
            If tipotienda = 1 Or tipotienda = 4 Then
               frmprecios.Show
            Else
               fnewprec.Show
            End If
        Case "F6"
             fcod.Show
    End Select
    'frabusqueda.Visible = False
    'StrTeclaPres = ""
    txtbusca.Text = ""
    gridbusca.Bookmark = reg
Exit Sub
Error:
MsgBox Err.Description
ElseIf KeyAscii <> 27 Then
   txtbusca.SetFocus
   txtbusca.Text = Chr(KeyAscii)
   txtbusca.SelStart = Len(txtbusca.Text)
End If
End Sub

'Private Sub ImgTapiz_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If Button = 2 Then
'   PopupMenu Me.MnuPri
'End If
'End Sub

Private Sub MnuInv_Click()
cmdBotones_Click 4
End Sub

Private Sub MnuCerCon_Click()
frmLogin.Show
frmLogin.chkPortatil.Visible = True
frmLogin.cmbServer.Visible = True
Unload frmAreaRecibo
End Sub

Private Sub mnuCerr_Click()
  End
End Sub

Private Sub mnuDep_Click()
 fdeptos.Show 1
End Sub

Private Sub Mnudespensa_Click()
  frmdespensa.Show
End Sub

Private Sub mnuenvios_Click()
cmdBotones_Click 1
End Sub

Private Sub mnufacturas_Click()
cmdBotones_Click 10
End Sub

Private Sub mnuFam_Click()
  ffamilia.Show 1
End Sub



Private Sub mnuLin_Click()
  flineas.Show 1
End Sub

Private Sub mnupedaba_Click()
cmdBotones_Click 7
End Sub

Private Sub mnupedprov_Click()
cmdBotones_Click 5
End Sub

Private Sub MnuPedsug_Click()
cmdBotones_Click 0
End Sub

Private Sub mnuProv_Click()
    lpprov = False
    fprov.Show 1
End Sub

Private Sub mnurecibo_Click()
cmdBotones_Click 1
End Sub

Private Sub MnuTraEnvNvo_Click()
 nOp = 1
 frmtraslados.Show
End Sub

Private Sub mnuSalAcerca_Click()
   frmAbout.Show 1
End Sub

Private Sub mnuSalCerr_Click()
   frmLogin.Show
   Unload Me
End Sub

Private Sub mnuSalFin_Click()
  cn.Close
  'MsgBox "RECUERDE QUE AL FINALIZAR LABORES DEBE EJECUTAR LA OPCION" & Chr(13) & "CORTE DE INVENTARIO DIARIO", vbInformation
  End
End Sub

Private Sub MnuTraEnv_Click()
  frmTrasladaEnv.Show
End Sub

Private Sub MnuTraRec_Click()
   frmTrasladaRec.Show
End Sub

Private Sub MnuPed_Click()
  frmpedidos.Show
End Sub

Private Sub MnuSes_Click()
frmLogin.Show
Unload frmAreaRecibo
End Sub

Private Sub MnuSolAct_Click()
SoloAct = Not MnuSolAct.Checked
MnuSolAct.Checked = SoloAct
cn.Execute "UPDATE usuarios SET SoloAct = " & IIf(SoloAct, 1, 0) & " WHERE clave ='" & Trim(Mid(cCveDesUsu, 1, 3)) & "'"
adobus.CursorType = adOpenKeyset
adobus.LockType = adLockOptimistic
adobus.CommandType = adCmdText
adobus.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
activo = IIf(SoloAct, " AND activo = 1", "")
If Nivel = "I" Then   'Si es usuario de compras internas
   adobus.RecordSource = "SELECT fecact,consec,activo,claprove,descripc,nomcorto,ltrim(STR(PAQUETES)) + ' X ' + STR(CONTENID,8,3)+ ' ' + MEDIDA AS PRESENTACION,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja,OFERTADO" & _
                         " FROM tfproduc WHERE interno = 1 ORDER BY descripc"
ElseIf Sql Then
   adobus.RecordSource = "SELECT precio1, fecact,consec,activo,claprove,descripc,nomcorto,ltrim(STR(PAQUETES)) + ' X ' + STR(CONTENID,8,3)+ ' ' + MEDIDA AS PRESENTACION,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja,OFERTADO" & _
                         " FROM tfproduc,preprod WHERE consec *= preclave and interno = 0 " & activo & " ORDER BY descripc"
Else
   ctfp = IIf(SoloAct, "tfproduc", "tfproduc")
        adobus.RecordSource = "SELECT fecact,consec,activo,claprove,descripc,nomcorto,ltrim(STR(PAQUETES)) + ' X ' + STR(CONTENID)+ ' ' + MEDIDA AS PRESENTACION,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja,OFERTADO" & _
                              " FROM " & ctfp & " WHERE interno = 0 " & activo & " ORDER BY descripc"
End If
'MsgBox adobus.RecordSource
adobus.Refresh
End Sub

Private Sub MnuTapiz_Click()
On Error GoTo Error:
   cmdlg.DialogTitle = "Seleccionar archivo para establecer el tapiz de la aplicacion"
   cmdlg.Filter = "Archivos de mapa de bits(*.bmp)|*.bmp|Archivos jpg(*.jpg)|*.jpg|Archivos jpeg(*.jpeg)|(*.jpeg)|Imagenes gif(*.gif)|*.gif|iconos (*.ico)|*.ico"
   cmdlg.CancelError = True
   cmdlg.ShowOpen
   If Trim(cmdlg.FileName) <> "" Then
       cn.Execute "UPDATE usuarios SET tapiz = '" & cmdlg.FileName & "' WHERE clave ='" & Trim(Mid(cCveDesUsu, 1, 3)) & "'"
       Me.Imgtapiz.Picture = LoadPicture(cmdlg.FileName)
   End If
   Exit Sub
   
Error:
  If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
     MsgBox Err.Description
  Else
     cn.Execute "UPDATE usuarios SET tapiz = '' WHERE clave ='" & Trim(Mid(cCveDesUsu, 1, 3)) & "'"
  End If

End Sub

Private Sub MnuTrasl_Click()
cmdBotones_Click 1
End Sub

Private Sub mnutdas_Click()
'SE IMPRIMEN LAS ETIQUETAS PARA EL INVENTARIO
RESP = InputBox("Escriba lo que quiere : CPU, MONI,TEC,MOUSE,IMP,OTRO")
CUA = InputBox("Cuantos")
inicial = InputBox("Numero Inicial ")
If Not verImpresora Then Exit Sub
nAncho = 250    'En puntos
Printer.ScaleMode = vbPoints
Printer.CurrentX = 10
espacio = "   "
For i = 1 To CUA
    Printer.Font = "arial"
    Printer.FontSize = "8"
    Printer.Print espacio & Trim(RESP)
    Printer.Font = "ZB 39* 10mil/2:1"
    Printer.FontSize = 40
    Printer.CurrentX = 10
    Printer.Print "  " & RESP & (Val(inicial) + i)
    Printer.EndDoc
Next
End Sub

Private Sub mnuusu_Click()
cmdlg.DialogTitle = "Abrir archivos Palm Procter"
cmdlg.FileName = App.Path & "\" & "Order_headers.txt"
cmdlg.Filter = "Archivos Planos Procter (Order_headers.txt) | Order_headers.txt"
cmdlg.ShowOpen


End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdConAceptar_Click
End If
End Sub

Private Sub ActPiso()
    Daocambio.Connect = "dbase III;"
    Daocambio.DatabaseName = "\\" & SERVIDOR & "\disco-c\vliventa"
    Daocambio.RecordsetType = vbRSTypeDynaset
    Daocambio.RecordSource = "cambios"
    Daocambio.Refresh

    'Cargo el tfproduc de la RED
    cmen = stbmensajes.Panels(1).Text
    stbmensajes.Panels(1).Text = "Procesando cambios para el area de Piso"
    DaoProdDbf.Connect = "dbase III;"
    DaoProdDbf.DatabaseName = "\\" & SERVIDOR & "\disco-c\vliventa"
    DaoProdDbf.RecordsetType = vbRSTypeDynaset
    DaoProdDbf.RecordSource = "SELECT * FROM tfproduc ORDER BY consec"
    DaoProdDbf.Refresh
    lblProd.Visible = True
    While Not Daocambio.Recordset.EOF
        nreg = nreg + 1: lblProd.Refresh
        CONSEC = Daocambio.Recordset!CONSEC + 1000000
        lblProd.Caption = "Productos procesados: " & CStr(nreg) & " --> " & CONSEC
        DaoProdDbf.Recordset.FindFirst "CONSEC = " & CONSEC
        If DaoProdDbf.Recordset.NoMatch Then
            DaoProdDbf.Recordset.AddNew
        Else
            DaoProdDbf.Recordset.Edit
        End If
        DaoProdDbf.Recordset!claprove = Daocambio.Recordset!claprove
        DaoProdDbf.Recordset!clafamil = Daocambio.Recordset!clafamil
        DaoProdDbf.Recordset!descripc = Trim(Daocambio.Recordset!descripc)
        DaoProdDbf.Recordset!NOMCORTO = Trim(Daocambio.Recordset!NOMCORTO)
        DaoProdDbf.Recordset!CONTENID = Daocambio.Recordset!CONTENID
        DaoProdDbf.Recordset!fletex = Daocambio.Recordset!fletex
        DaoProdDbf.Recordset!flesub = Daocambio.Recordset!flesub
        DaoProdDbf.Recordset!medida = Daocambio.Recordset!medida
        DaoProdDbf.Recordset!PAQUETES = Daocambio.Recordset!PAQUETES
        DaoProdDbf.Recordset!descto01 = Daocambio.Recordset!descto01
        DaoProdDbf.Recordset!descto02 = Daocambio.Recordset!descto02
        DaoProdDbf.Recordset!descto03 = Daocambio.Recordset!descto03
        DaoProdDbf.Recordset!descto04 = Daocambio.Recordset!descto04
        DaoProdDbf.Recordset!descto05 = Daocambio.Recordset!descto05
        DaoProdDbf.Recordset!descto06 = Daocambio.Recordset!descto06
        DaoProdDbf.Recordset!descefec = Daocambio.Recordset!descefec
        DaoProdDbf.Recordset!porcargo = Daocambio.Recordset!porcargo
        DaoProdDbf.Recordset!fletes = Daocambio.Recordset!fletes
        DaoProdDbf.Recordset!otrosrec = Daocambio.Recordset!otrosrec
        DaoProdDbf.Recordset!PRELISTA = Daocambio.Recordset!PRELISTA
        DaoProdDbf.Recordset!ieps = Daocambio.Recordset!ieps
        DaoProdDbf.Recordset!tasaieps = Daocambio.Recordset!tasaieps
        DaoProdDbf.Recordset!iva = Daocambio.Recordset!iva
        DaoProdDbf.Recordset!prepaque = Daocambio.Recordset!prepaque
        DaoProdDbf.Recordset!PRECAJA = Daocambio.Recordset!PRECAJA
        DaoProdDbf.Recordset!prelib1 = Daocambio.Recordset!prelib1
        DaoProdDbf.Recordset!prelib2 = Daocambio.Recordset!prelib2
        DaoProdDbf.Recordset!ganancaj = Daocambio.Recordset!ganancaj
        DaoProdDbf.Recordset!gananpaq = Daocambio.Recordset!gananpaq
        DaoProdDbf.Recordset!gananlib2 = Daocambio.Recordset!gananlib1
        DaoProdDbf.Recordset!tantos = Daocambio.Recordset!tantos
        DaoProdDbf.Recordset!entre = Daocambio.Recordset!entre
        DaoProdDbf.Recordset!proceden = Daocambio.Recordset!proceden
        DaoProdDbf.Recordset!CONSEC = CONSEC
        DaoProdDbf.Recordset!oferta = Daocambio.Recordset!oferta
        DaoProdDbf.Recordset!pedir = IIf(Daocambio.Recordset!pedir = 0, 1, 0)
        DaoProdDbf.Recordset!fecact = Daocambio.Recordset!fecact
        DaoProdDbf.Recordset!tasaiva = Daocambio.Recordset!tasaiva
        DaoProdDbf.Recordset!peso = Daocambio.Recordset!peso
        DaoProdDbf.Recordset!COSTOPAQ = Daocambio.Recordset!COSTOPAQ
        DaoProdDbf.Recordset!costocaj = Daocambio.Recordset!costocaj
        DaoProdDbf.Recordset!barraspza = Daocambio.Recordset!barraspza
        DaoProdDbf.Recordset!barrascaja = Daocambio.Recordset!barrascaja
        DaoProdDbf.Recordset!claveprov = Trim(Daocambio.Recordset!claveprov)
        DaoProdDbf.Recordset.Update
        Daocambio.Recordset.MoveNext
    Wend
    Daocambio.Recordset.Close
    Set Daocambio.Recordset = Nothing
    stbmensajes.Panels(1).Text = cmen
End Sub

Sub ActBodega()
Dim cnAct As ADODB.Connection
'On Error Resume Next

cMens = stbmensajes.Panels(1).Text
stbmensajes.Panels(1).Text = "Leyendo archivo de productos para cambio de precios"
lblProd.Visible = True

stbmensajes.Panels(1).Text = "Procesando cambios para el area de Bodega"
cn.Execute "DELETE FROM cambios"
 nreg = 0
 Adodbf.CommandType = adCmdText
 Adodbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & "\\" & SERVIDOR & "\disco-c\paso"
 Adodbf.RecordSource = "SELECT * FROM cambios"
 Adodbf.Refresh

While Not Adodbf.Recordset.EOF
  nreg = nreg + 1: lblProd.Refresh
  CONSEC = Adodbf.Recordset!CONSEC
  lblProd.Caption = "Productos procesados: " & CStr(nreg) & " --> " & CONSEC
  nprecio = Adodbf.Recordset!costocaj
  If IsNull(Adodbf.Recordset!descripc) Then
     descr = ""
  Else
     descr = Replace(Adodbf.Recordset!descripc, "'", " ")
  End If
  If IsNull(Adodbf.Recordset!NOMCORTO) Then
     NOMCORTO = ""
  Else
     NOMCORTO = Replace(Adodbf.Recordset!NOMCORTO, "'", " ")
  End If
  
  cn.Execute "INSERT INTO CAMBIOS(CLAPROVE,CLAFAMIL,descripc,nomcorto,contenid,fletex,flesub,medida,paquetes,usuario,descto01,descto02,descto03,descto04,descto05,descto06,descefec,porcargo,fletes,Otrosrec,maniobras,prelista,ieps,tasaieps,iva,prepaque,precaja,prelib1,prelib2,prelib3,prelib4," & _
             "Ganancaj,GananPaq,Gananlib1,GananLib2,Tantos,entre,proceden,consec,oferta,pedir,fecact,tasaiva,peso,costopaq,costocaj,barraspza,barrascaja,precosto) VALUES('" & Adodbf.Recordset!claprove & "','" & Adodbf.Recordset!clafamil & "','" & descr & "','" & NOMCORTO & _
             "'," & Adodbf.Recordset!CONTENID & "," & IIf(IsNull(Adodbf.Recordset!fletex), 0, Adodbf.Recordset!fletex) & "," & IIf(IsNull(Adodbf.Recordset!flesub), 0, Adodbf.Recordset!flesub) & ",'" & Adodbf.Recordset!medida & "'," & Adodbf.Recordset!PAQUETES & ",'" & Adodbf.Recordset!USUARIO & "'," & Adodbf.Recordset!descto01 & "," & Adodbf.Recordset!descto02 & "," & Adodbf.Recordset!descto03 & "," & Adodbf.Recordset!descto04 & "," & Adodbf.Recordset!descto05 & "," & Adodbf.Recordset!descto06 & "," & Adodbf.Recordset!descefec & "," & Adodbf.Recordset!porcargo & "," & IIf(IsNull(Adodbf.Recordset!fletes), 0, Adodbf.Recordset!fletes) & "," & Adodbf.Recordset!otrosrec & "," & IIf(IsNull(Adodbf.Recordset!maniobras), 0, Adodbf.Recordset!maniobras) & "," & Adodbf.Recordset!PRELISTA & "," & Adodbf.Recordset!ieps & "," & _
             Adodbf.Recordset!tasaieps & "," & Adodbf.Recordset!iva & "," & Adodbf.Recordset!prepaque & "," & Adodbf.Recordset!PRECAJA & "," & Adodbf.Recordset!prelib1 & "," & Adodbf.Recordset!prelib2 & "," & Adodbf.Recordset!prelib3 & "," & Adodbf.Recordset!prelib4 & "," & Adodbf.Recordset!ganancaj & "," & Adodbf.Recordset!gananpaq & "," & Adodbf.Recordset!gananlib1 & "," & Adodbf.Recordset!gananlib2 & "," & Adodbf.Recordset!tantos & "," & Adodbf.Recordset!entre & "," & IIf(IsNull(Adodbf.Recordset!proceden), 0, Adodbf.Recordset!proceden) & "," & Adodbf.Recordset!CONSEC & "," & Adodbf.Recordset!oferta & "," & _
             IIf(Adodbf.Recordset!pedir = 1, 0, 1) & ",'" & Adodbf.Recordset!fecact & "'," & IIf(IsNull(Adodbf.Recordset!tasaiva), 0, Adodbf.Recordset!tasaiva) & "," & IIf(IsNull(Adodbf.Recordset!peso), 0, Adodbf.Recordset!peso) & _
             "," & IIf(IsNull(Adodbf.Recordset!COSTOPAQ), 0, Adodbf.Recordset!COSTOPAQ) & "," & IIf(IsNull(Adodbf.Recordset!costocaj), 0, Adodbf.Recordset!costocaj) & "," & IIf(IsNull(Adodbf.Recordset!barraspza), 0, Adodbf.Recordset!barraspza) & "," & IIf(IsNull(Adodbf.Recordset!barrascaja), 0, Adodbf.Recordset!barrascaja) & "," & IIf(IsNull(nprecio), 0, nprecio) & ")"
  Adodbf.Recordset.MoveNext
Wend
'A puerto escondido se le agrega el 2%
If Trim(Mid(cSucursal, 1, 2)) = "55" Then
   cn.Execute "UPDATE cambios SET Prelib2 = ROUND( prelib2  + (prelib2 * 0.02) ,1,-2) + .10 FROM tfproduc t WHERE t.consec = cambios.consec + 1000000 And NOT t.descripc like '%CIGARRO%' AND Not cambios.claprove in ('C66','C98','P17','P12','L34','C73')"
   cn.Execute "UPDATE cambios SET prelib1 = prelib2"   'El precio de bodega e intermedio es el mismo
   cn.Execute "UPDATE cambios SET precaja = prelib2 WHERE prelib2 > precaja"  'Si es mas caro el de bodega que el envío"
   'cn.Execute "UPDATE cambios SET prepaque = ROUND( prepaque  + (prepaque * 0.02) ,1,-2) + .10, precaja = ROUND( precaja  + (precaja * 0.02) ,1,-2) + .10, Prelib1 = ROUND( prelib1  + (prelib1 * 0.02) ,1,-2) + .10, Prelib2 = ROUND( prelib2  + (prelib2 * 0.02) ,1,-2) + .10 , Prelib3 = ROUND( prelib3  + (prelib3 * 0.02) ,1,-2) + .10 FROM tfproduc t WHERE t.consec = cambios.consec + 1000000 And NOT t.descripc like '%CIGARRO%' AND NOT T.claprove IN ('C66','C98','P17','P12','L34','C73')"
   'cn.Execute "UPDATE cambios SET prepaque = ROUND( prepaque  + (prepaque * 0.01) ,1,-2) + .10, precaja = ROUND( precaja  + (precaja * 0.01) ,1,-2) + .10, Prelib1 = ROUND( prelib1  + (prelib1 * 0.01) ,1,-2) + .10, Prelib2 = ROUND( prelib2  + (prelib2 * 0.01) ,1,-2) + .10 , Prelib3 = ROUND( prelib3  + (prelib3 * 0.01) ,1,-2) + .10 FROM tfproduc t WHERE t.consec = cambios.consec + 1000000 And t.descripc like '%CIGARRO%'"
ElseIf Trim(Mid(cSucursal, 1, 2)) = "28" Then      'ISTMO
   'cn.Execute "UPDATE cambios SET precaja = prelib1"   EL INTERDIO ES IGUAL AL DE ENVIO
   cn.Execute "UPDATE cambios SET prepaque = ROUND( prepaque  + (prepaque * 0.01) ,1,-2) + .10, precaja = ROUND( precaja  + (precaja * 0.01) ,1,-2) + .10, Prelib1 = ROUND( prelib1  + (prelib1 * 0.01) ,1,-2) + .10, Prelib2 = ROUND( prelib2  + (prelib2 * 0.01) ,1,-2) + .10 , Prelib3 = ROUND( prelib3  + (prelib3 * 0.01) ,1,-2) + .10, Prelib4 = ROUND( prelib4  + (prelib4 * 0.01) ,1,-2) + .10 FROM tfproduc t WHERE t.consec = cambios.consec + 1000000 And NOT t.descripc like '%CIGARRO%'"
   'cn.Execute "UPDATE cambios SET prepaque = ROUND( prepaque  + (prepaque * 0.01) ,1,-2) + .10, precaja = ROUND( precaja  + (precaja * 0.02) ,1,-2) + .10, Prelib1 = ROUND( prelib1  + (prelib1 * 0.02) ,1,-2) + .10, Prelib2 = ROUND( prelib2  + (prelib2 * 0.02) ,1,-2) + .10 , Prelib3 = ROUND( prelib3  + (prelib3 * 0.02) ,1,-2) + .10 FROM tfproduc t WHERE t.consec = cambios.consec + 1000000 And ( t.descripc like '%FRIJOL NEGRO BUENO%')"
End If
Set cnAct = New ADODB.Connection
cnAct.ConnectionTimeout = 0
cnAct.CommandTimeout = 0
cnAct.Open cCadConex
'ANIMA
stbmensajes.Panels(1).Text = "Verificando productos nuevos en catálogo de productos"
cnAct.Execute "INSERT INTO TFPRODUC (CONSEC,DESCRIPC,NOMCORTO,CONTENID,FLETEX,FLESUB,MEDIDA,PAQUETES,COSTOPAQ,COSTOCAJ,PESO,BARRASPZA,BARRASCAJA,TASAIEPS,CLAPROVE,PROCEDENCIA,FECACT,CAJAS,ENCAJAS,PRECOSTO,CLAFAMIL,ACTIVO,OFERTADO,LINEA,IVA,IEPS) " & _
           "SELECT PROD = cast(CONSEC as decimal )+ 1000000,DESCRIPC,NOMCORTO,CONTENID,FLETEX,FLESUB,MEDIDA,PAQUETES,COSTOPAQ,PRELISTA,PESO,BARRASPZA,BARRASCAJA,TASAIEPS,CLAPROVE,PROCEDEN,FECACT,TANTOS,ENTRE,COSTOCAJ,CLAFAMIL,PEDIR,OFERTA,LINEA = 757,IVA,IEPS FROM CAMBIOS " & _
           "WHERE NOT cast(CONSEC as decimal ) + 1000000 IN (SELECT CAST(consec AS DECIMAL) FROM TFPRODUC)"

stbmensajes.Panels(1).Text = "Actualizando catálogo de productos"
cnAct.Execute "UPDATE TFPRODUC SET DESCRIPC = CAMBIOS.DESCRIPC,NOMCORTO = CAMBIOS.NOMCORTO, CONTENID=CAMBIOS.CONTENID, FLETEX=CAMBIOS.FLETEX, FLESUB=CAMBIOS.FLESUB, MEDIDA=CAMBIOS.MEDIDA, PAQUETES=CAMBIOS.PAQUETES, COSTOPAQ=CAMBIOS.COSTOPAQ, COSTOCAJ=CAMBIOS.PRELISTA, " & _
           "BARRASPZA=CAMBIOS.BARRASPZA, BARRASCAJA=CAMBIOS.BARRASCAJA, TASAIEPS=CAMBIOS.TASAIEPS, CLAPROVE=CAMBIOS.CLAPROVE, PROCEDENCIA=CAMBIOS.PROCEDEN, FECACT=CAMBIOS.FECACT, CAJAS=CAMBIOS.TANTOS, ENCAJAS=CAMBIOS.ENTRE, " & _
           "OFERTADO=CAMBIOS.OFERTA, ACTIVO=CAMBIOS.PEDIR, PRECOSTO=CAMBIOS.COSTOCAJ, CLAFAMIL=CAMBIOS.CLAFAMIL, IVA=CAMBIOS.IVA, IEPS=CAMBIOS.IEPS FROM CAMBIOS WHERE CAST(TFPRODUC.CONSEC AS DECIMAL) = CAMBIOS.CONSEC + 1000000"

stbmensajes.Panels(1).Text = "Verificando productos nuevos en catálogo de precios"
cnAct.Execute "INSERT INTO PREPROD (PRECLAVE,PRECIO1,PRECIO2,PRECIO3,PRECIO4,PRECIO5,PRECIO6,FECHAACT,PRUSUARIO) " & _
             "SELECT PROD = cast(CONSEC as decimal )+ 1000000, PREPAQUE,PRECAJA,PRELIB1,PRELIB2,PRELIB3,PRELIB4,FECACT,Usuario FROM CAMBIOS " & _
             "WHERE NOT cast(CONSEC as decimal ) + 1000000 IN (SELECT CAST(PRECLAVE AS DECIMAL) FROM PREPROD)"

stbmensajes.Panels(1).Text = "Actualizando catálogo de precios"
cnAct.Execute "UPDATE preprod SET precio1=prepaque, precio2=precaja, precio3=prelib1, precio4=prelib2, precio5 = prelib3, precio6 = prelib4, fechaact=cambios.fecact, prusuario = usuario FROM CAMBIOS " & _
           "WHERE CAST(PRECLAVE AS DECIMAL) = CAMBIOS.CONSEC + 1000000"

stbmensajes.Panels(1).Text = "Verificando productos nuevos en catálogo de cargos"
cnAct.Execute "INSERT INTO cargos(caprod,cargo1,cargo_efectivo,iva,ieps) " & _
             "SELECT PROD = cast(CONSEC as decimal )+ 1000000, porcargo,otrosrec,Iva,Ieps FROM CAMBIOS " & _
             "WHERE NOT cast(CONSEC as decimal ) + 1000000 IN (SELECT CAST(caprod AS DECIMAL) FROM cargos)"

stbmensajes.Panels(1).Text = "Actualizando catálogo de cargos"
cnAct.Execute "UPDATE cargos SET maniobras = cambios.maniobras, flete_efectivo = fletes, cargo1 = porcargo, cargo_efectivo = otrosrec, cargos.iva = cambios.iva, cargos.ieps = cambios.ieps FROM CAMBIOS " & _
           "WHERE CAST(caprod AS DECIMAL) = CAMBIOS.CONSEC + 1000000"

stbmensajes.Panels(1).Text = "Verificando productos nuevos en catálogo de descuentos"
cnAct.Execute "INSERT INTO descuentos(deprod,decto1,decto2,decto3,dectoOferta,decto5,dectoFinanciero,dectoefectivo) " & _
             "SELECT PROD = cast(CONSEC as decimal )+ 1000000, descto01,descto02,descto03,descto04,descto05,descto06,descefec FROM CAMBIOS " & _
             "WHERE NOT cast(CONSEC as decimal ) + 1000000 IN (SELECT CAST(deprod AS DECIMAL) FROM descuentos)"

stbmensajes.Panels(1).Text = "Actualizando catálogo de descuentos"
cnAct.Execute "UPDATE descuentos SET decto1 = descto01, decto2= descto02, decto3= descto03, dectoOferta= descto04, decto5= descto05, dectoFinanciero= descto06, dectoEfectivo = descefec FROM CAMBIOS " & _
           "WHERE CAST(deprod AS DECIMAL) = CAMBIOS.CONSEC + 1000000"

stbmensajes.Panels(1).Text = cMens
'AdoDbf.Recordset.Close
'Set AdoDbf.Recordset = Nothing
End Sub

Private Sub ActProv()
cmen = Me.stbmensajes.Panels(1).Text
stbmensajes.Panels(1).Text = "Actualizando catálogo de proveedores"
'Me.stbMensajes.SimpleText = Space(30) & "Espere Actualizando Proveedores..."
'adodbf.Connect = "dbase III;"
'adodbf.DatabaseName = "\\" & SERVIDOR & "\disco-c\paso"
'adodbf.RecordsetType = Table
'adodbf.RecordSource = "catprov"
'adodbf.Refresh

Adodbf.CommandType = adCmdText
Adodbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & "\\" & SERVIDOR & "\disco-c\paso"
Adodbf.RecordSource = "SELECT * FROM catprov"
Adodbf.Refresh

AdoProv.CommandType = adCmdText
AdoProv.ConnectionString = cCadConex '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
AdoProv.RecordSource = "select * from catprov"
AdoProv.Refresh
clave = ""
If Adodbf.Recordset.RecordCount > 0 Then
    probar1.Min = 0
    probar1.Max = Adodbf.Recordset.RecordCount
    Lbltrans.Visible = True
    Lbltrans.Refresh
    probar1.Visible = True
    v = 0
    Adodbf.Recordset.MoveFirst
    'AdoProv.Refresh
    While Not Adodbf.Recordset.EOF
      Me.lblProd.Caption = "Proveedor: " & Str(v)
      lblProd.Refresh
      v = v + 1
      probar1.Value = v
      clave = Adodbf.Recordset!prove
      concomilla = InStr(1, Adodbf.Recordset!NOMPROVE, "'")
      If concomilla > 0 Then
         xnomprove = Mid(Adodbf.Recordset!NOMPROVE, 1, concomilla - 1) + "  " + Mid(Adodbf.Recordset!NOMPROVE, concomilla + 1, Len(Adodbf.Recordset!NOMPROVE))
      Else
         xnomprove = Adodbf.Recordset!NOMPROVE
      End If
      xnomprove = Mid(xnomprove, 1, 150)
      If Not (AdoProv.Recordset.BOF And AdoProv.Recordset.EOF) Then AdoProv.Recordset.MoveFirst
      AdoProv.Recordset.Find "prove = '" & Trim(clave) & "'"
      If AdoProv.Recordset.EOF Then
         CADENA = "INSERT INTO CATPROV (prove,nomprove,tipo,ACTIVO,BACKORDER) values ('" & _
         IIf(Not IsNull(Adodbf.Recordset!prove), Adodbf.Recordset!prove, 0) & "','" & IIf(Not IsNull(Adodbf.Recordset!NOMPROVE), xnomprove, 0) & "'," & _
         "'I',1,1)"
         cn.Execute CADENA
      Else
         If IsNull(Adodbf.Recordset!telpro) Then
            TELE = ""
         Else
            TELE = Replace(Adodbf.Recordset!telpro, "'", "")
         End If
         'En bodega carbonera si se debe actualizar el tipo de proveedor
         If tipotienda = 2 Then actTipo = ", tipo = '" & Adodbf.Recordset!tipo & "'"
         CADENA = "UPDATE catprov SET nomprove = '" & IIf(Not IsNull(Mid(Adodbf.Recordset!NOMPROVE, 1, 150)), xnomprove, 0) & "', dirpro = '" & _
         IIf(Not IsNull(Mid(Adodbf.Recordset!dirpro, 1, 50)), Mid(Adodbf.Recordset!dirpro, 1, 50), 0) & "', colpro = '" & IIf(Not IsNull(Adodbf.Recordset!colpro), Adodbf.Recordset!colpro, 0) & "', delpro =  '" & _
         IIf(Not IsNull(Adodbf.Recordset!delpro), Adodbf.Recordset!delpro, 0) & "', codpro =  '" & IIf(Not IsNull(Adodbf.Recordset!codpro), Adodbf.Recordset!codpro, 0) & "', ciupro = '" & _
         IIf(Not IsNull(Adodbf.Recordset!ciupro), Adodbf.Recordset!ciupro, 0) & "', locpro = '" & IIf(Not IsNull(Adodbf.Recordset!LOCPRO), Adodbf.Recordset!LOCPRO, 0) & "', telpro = '" & _
          TELE & "'" & actTipo & " WHERE prove = '" & Trim(clave) & "'"
         'MsgBox CADENA
         cn.Execute CADENA
      End If
      Adodbf.Recordset.MoveNext
     Wend
End If
probar1.Visible = False
Lbltrans.Visible = False
Lbltrans.Refresh
Cmdopt(1).SetFocus
stbmensajes.Panels(1).Text = cmen
End Sub


Function CalCosto() As Double
Dim nprecio As Double
    nprecio = 0
    nprecio = IIf(Not IsNull(DaoProdDbf.Recordset!PRELISTA), DaoProdDbf.Recordset!PRELISTA, 0)
    'calcula cargos %
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!iva), DaoProdDbf.Recordset!iva, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!ieps), DaoProdDbf.Recordset!ieps, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio + (nprecio * (npreciopaso / 100))
    End If
    'nprecio = Round(nprecio, 2)
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto01), DaoProdDbf.Recordset!descto01, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio - (nprecio * (npreciopaso / 100))
    End If
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto02), DaoProdDbf.Recordset!descto02, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio - (nprecio * (npreciopaso / 100))
    End If
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto03), DaoProdDbf.Recordset!descto03, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio - (nprecio * (npreciopaso / 100))
    End If
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto04), DaoProdDbf.Recordset!descto04, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio - (nprecio * (npreciopaso / 100))
    End If
    
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto05), DaoProdDbf.Recordset!descto06, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio - (nprecio * (npreciopaso / 100))
    End If
        
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descto06), DaoProdDbf.Recordset!descto06, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio - (nprecio * (npreciopaso / 100))
    End If
    
    'descuento efectivo
    npreciopaso = IIf(Not IsNull(DaoProdDbf.Recordset!descefec), DaoProdDbf.Recordset!descefec, 0)
    If npreciopaso > 0 Then
       nprecio = nprecio - npreciopaso
    End If
    'nprecio = Round(nprecio, 2)
    CalCosto = nprecio
End Function



Sub ANIMA()
On Error Resume Next
ani1.AutoPlay = True
'Ani1.Open ("C:\PITIC\PROGRAMAS\GRAPHICS\AVIS\FILEMOVE.AVI")
ani1.Open (App.Path & "\FILEMOVE.AVI")
'SendKeys "{ENTER}"
'MsgBox "SD"
End Sub

Private Sub txtbusca_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
If KeyAscii = 13 Then
   'se cargan los productos hasta que realmente desean buscar un producto
   'ya que se cargaban con la forma aunque no se realizaran busquedas y se tardaba
   If consulta = "" Then
        adobus.CursorType = adOpenKeyset
        adobus.LockType = adLockOptimistic
        adobus.CommandType = adCmdText
        adobus.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
        activo = IIf(SoloAct, " AND activo = 1", "")
        If Nivel = "I" Then   'Si es usuario de compras internas
            adobus.RecordSource = "select fecact,consec,activo,claprove,descripc,nomcorto,ltrim(STR(PAQUETES)) + ' X ' + STR(CONTENID,8,3)+ ' ' + MEDIDA AS PRESENTACION,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja,OFERTADO" & _
            " FROM tfproduc WHERE interno = 1 ORDER BY descripc"
        Else
            If Sql Then
               adobus.RecordSource = "SELECT precio1, fecact,consec,activo,claprove,descripc,nomcorto,ltrim(STR(PAQUETES)) + ' X ' + STR(CONTENID,8,3)+ ' ' + MEDIDA AS PRESENTACION,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja,OFERTADO" & _
                                     " FROM tfproduc,preprod WHERE consec *= preclave and interno = 0 " & activo & " ORDER BY descripc"
            Else
               ctfp = IIf(SoloAct, "tfproduc", "tfproduc")
               adobus.RecordSource = "SELECT fecact,consec,activo,claprove,descripc,nomcorto,ltrim(STR(PAQUETES)) + ' X ' + STR(CONTENID)+ ' ' + MEDIDA AS PRESENTACION,fletex,flesub,medida,paquetes,procedencia,costopaq,costocaj,barraspza,barrascaja,OFERTADO" & _
                                     " FROM " & ctfp & " WHERE interno = 0 " & activo & " ORDER BY descripc"
            End If
        End If
        adobus.Refresh
    End If
    txtbusca.Text = UCase(Trim(txtbusca.Text))
    Select Case StrTeclaPres
        Case "F2"
            consulta = " descripc like '" & Trim(txtbusca.Text) & "*'"
        Case "F3"
            consulta = " descripc like '" & Trim(txtbusca.Text) & "*'"
        Case "F4"
            consulta = " barraspza = " & Val(Trim(txtbusca.Text))
        Case "F5"
            consulta = " consec ='" & Trim(txtbusca.Text) & "'"
    End Select
    adobus.Recordset.MoveFirst
    If consulta <> " descripc like '*'" Then
       adobus.Recordset.Find consulta
    End If
    If adobus.Recordset.EOF Then
       Select Case StrTeclaPres
         Case "F2", "F3"
             N = 1
             While adobus.Recordset.EOF
                consulta = " '" & Mid(Trim(txtbusca.Text), 1, Len(Trim(txtbusca.Text)) - N) & "%'"
                adobus.Recordset.MoveFirst
                If Trim(consulta) = "'%'" Then
                   MsgBox "NO EXISTE NINGUN PRODUCTO COINCIDENTE", vbInformation
                   Exit Sub
                End If
                adobus.Recordset.Find " descripc like " & consulta
                N = N + 1
             Wend
         Case "F4"
             consulta = " barraspza = " & Val(Trim(txtbusca.Text))
         Case "F5"
             consulta = " consec ='" & Trim(txtbusca.Text) & "'"
       End Select
       adobus.Recordset.MoveFirst
       adobus.Recordset.Find consulta
    End If
    gridbusca.Enabled = True
    gridbusca.Visible = True
    gridbusca.SetFocus
End If
Exit Sub
Error:
  Exit Sub
End Sub

Private Sub ActExiCdc()
On Error GoTo Error:
Dim nDias As Integer
Dim FolPed As String
Dim rsttemp As ADODB.Recordset
Dim cArch  As String
Dim cProv As String
Dim SUC As String
Dim cnFoxPro As ADODB.Connection
   cmdlg.FileName = ""
   cmdlg.CancelError = True
   cmdlg.DialogTitle = "Abrir archivo de existencias de CDC"
   cvesuc = Trim(Mid(cSucursal, 1, 2))
   If cvesuc = "16" Then
      cmdlg.Filter = "Archivo de Existencias Cdc (PITICO10.mdb) | Pitico10.mdb"
   Else
      cmdlg.Filter = "Archivo de Existencias MCABRERA(Exibod.dbf) | Exibod.dbf"
   End If
   cmdlg.ShowOpen
   cRutArc = cmdlg.FileName

   If Trim(cRutArc) = "" Or IsNull(cRutArc) Then
      MsgBox "DEBE ESPECIFICAR UN NOMBRE DE ARCHIVO", vbExclamation
      Exit Sub
   End If

   'Siempre el nomnbre de archivo es de 8 Caracteres
   For N = 1 To Len(cRutArc)
      If Mid(cRutArc, N, 1) = "\" Then nPos = N
   Next
   cruta = Mid(cRutArc, 1, nPos)
   cArch = Mid(cRutArc, nPos + 1)
   cArch = Mid(cArch, 1, Len(cArch) - 4) 'Le quito la extension porque si no marca error en la consula SQL.
    
   AdoProv.CursorType = adOpenKeyset
   If cvesuc = "16" Then
      AdoProv.ConnectionString = "DSN=PITICOMDB;DBQ=P:\PITICO\PITICO10\PITICO10.mdb;DefaultDir=P:\PITICO\PITICO10;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"
      AdoProv.RecordSource = "SELECT * FROM inventario"
   Else
      AdoProv.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & cruta
      AdoProv.RecordSource = "SELECT * FROM " & cArch
   End If
   AdoProv.Refresh
   
   If AdoProv.Recordset.BOF And AdoProv.Recordset.EOF Then
      MsgBox "EL ARCHIVO SELECCIONADO ESTA VACIO, NO EXISTEN DATOS PARA IMPORTAR", vbExclamation
      Exit Sub
   Else
      AdoProv.Recordset.MoveFirst
   End If
   ANIMA
   probar1.Min = 0
   probar1.Max = AdoProv.Recordset.RecordCount
   Lbltrans.Visible = True
   Lbltrans.Refresh
   lblProd.Visible = True
   probar1.Visible = True
   v = 0
   cn.Execute "UPDATE inventario SET incantcdc = 0 , Incantpzacdc = 0"
   While Not AdoProv.Recordset.EOF
      probar1.Value = v
      lblProd.Caption = "Productos actualizados: " & CStr(v): lblProd.Refresh
      If cvesuc = "16" Then
         cn.Execute "UPDATE INVENTARIO SET INCANTCDC = " & AdoProv.Recordset!InCant & ",INCANTPZACDC = " & AdoProv.Recordset!InCantPza & " WHERE INprod = '" & AdoProv.Recordset!Inprod & "'"
      Else
         cn.Execute "UPDATE INVENTARIO SET INCANTCDC = " & AdoProv.Recordset!exicaja & ",INCANTPZACDC = " & AdoProv.Recordset!exipza & " WHERE INprod = '" & AdoProv.Recordset!CONSEC & "'"
      End If
      v = v + 1
      AdoProv.Recordset.MoveNext
   Wend
   probar1.Visible = False
   Lbltrans.Visible = False
   Lbltrans.Refresh
   ani1.Close
   MsgBox "LA IMPORTACION DEL CENTRO DE DISTRIBUCION CARBONERA SE REALLIZO CORRECTAMENTE", vbInformation
Exit Sub
Error:
   If Err.Number <> 32755 Then  'Numero de error al presionar el boton cancelar
      MsgBox Err.Description
      ani1.Close
   End If

End Sub

Private Sub exportacambios()
'SE DEBEN MANDAR LOS DESCUENTOS,CARGOS, ESCALAS QUE SON APLICADOS A EL PRECIO y NO A LOS PEDIDOS
'SE DEBE RESPETAR EL RANGO DE FECHAS

AdoProd.CursorType = adOpenKeyset
AdoProd.LockType = adLockOptimistic
AdoProd.CommandType = adCmdText
AdoProd.ConnectionString = strconnect
txtinicio1 = DateAdd("d", -1, txtinicio.Text)
txtfinal1 = DateAdd("d", 1, txtfinal.Text)


RESP = MsgBox("DESEA GENERAR CAMBIOS CON PRODUCTOS INACTIVOS ?", vbYesNoCancel + vbQuestion)
If RESP = vbYes Then
    CADENA = "select * from tfproduc,descuentos,preprod,cargos,margen  where consec = deprod and consec=margen.producto and consec = preclave  and caprod = consec " & _
    " and fecact >  '" & txtinicio1 & " ' and fecact < '" & txtfinal1 & " ' " & compInt
ElseIf RESP = vbCancel Then
    Exit Sub
    CADENA = "select * from tfproduc,descuentos,preciostemp t,cargos,margen  where consec = deprod and consec=margen.producto and consec = t.producto  and caprod = consec  AND ACTIVO = 1  " & _
     compInt
Else
    CADENA = "select * from tfproduc,descuentos,preprod,cargos,margen  where consec = deprod and consec=margen.producto and consec = preclave  and caprod = consec  AND ACTIVO = 1  " & _
    " and fecact >  '" & txtinicio1 & " ' and fecact < '" & txtfinal1 & " ' " & compInt
End If
AdoProd.RecordSource = CADENA
AdoProd.Refresh
MsgBox "Total de Productos a Generar : " & AdoProd.Recordset.RecordCount, vbInformation
Exporta ("CAMBIOS")
End Sub

Sub Exporta(opcion)
'On Error GoTo error:
'este proceso utiliza 3 dbfs en el directorio
'prodnob.dbf es la estructura de productos que se copia al archivo producto
'producto.dbf y productox.dbf se utiliza para que el recordset libere la tabla de productos
'y se pueda utilizar en el proceso de importacion o a la inversa
'POR NORMA TODO LOS DESCUENTOS DEBEN SER CON BASE A LOS QUE SE
'APLICARON AL PRECIO
AdoProv.CursorType = adOpenKeyset
AdoProv.LockType = adLockOptimistic
AdoProv.CommandType = adCmdText
AdoProv.ConnectionString = strconnect '"DSN=PITICO;SERVER=DATUM;PWD=TIJERAS ;UID=dba"
AdoProv.RecordSource = "select * from catprov"
AdoProv.Refresh

Dim v As Integer
Dim fs
If AdoProd.Recordset.RecordCount > 0 Then
    Set fs = CreateObject("Scripting.FileSystemObject")
    ' DEBE ESTAR CONECTADO A LA UNIDAD P :
    If Nivel = "I" Then   'Compras internas
       fs.copyfile "P:\reportes\prodnob.dbf", "p:\pasoInt\producto.dbf", True
       fs.copyfile "p:\reportes\provnob.dbf", "P:\pasoInt\catprov.dbf", True
       rutacambios = "p:\pasoInt"
    Else
       fs.copyfile "P:\reportes\prodnob.dbf", "p:\paso\producto.dbf", True
       fs.copyfile "p:\reportes\provnob.dbf", "P:\paso\catprov.dbf", True
       'fs.copyfile "p:\reportes\promnob.dbf", "P:\paso\prom.dbf", True
       rutacambios = "p:\paso"
    End If
    'Adodbf.Connect = "dbase III;"
    'Adodbf.DatabaseName = rutacambios
    'Adodbf.RecordsetType = Table
    'Adodbf.RecordSource = "producto"
    'Adodbf.Refresh
    
    Adodbf.CommandType = adCmdText
    Adodbf.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PITICODBF;SourceDB=" & rutacambios
    Adodbf.RecordSource = "SELECT * FROM producto"
    Adodbf.Refresh
    
    probar1o.Min = 0
    probar1o.Max = IIf(AdoProd.Recordset.RecordCount > 0, AdoProd.Recordset.RecordCount, 1)
    lbltransO.Visible = True
    lbltransO.Refresh
    probar1o.Visible = True
    v = 0
    AdoProd.Recordset.MoveFirst
    suma = 0
    While Not AdoProd.Recordset.EOF
        v = v + 1
        probar1o.Value = v
        concomilla = InStr(1, AdoProd.Recordset!descripc, "'")
        If concomilla > 0 Then
           XDESCRIPC = Mid(AdoProd.Recordset!descripc, 1, concomilla - 1) + "  " + Mid(AdoProd.Recordset!descripc, concomilla + 1, Len(AdoProd.Recordset!descripc))
        Else
           XDESCRIPC = AdoProd.Recordset!descripc
        End If
        concomillan = InStr(1, AdoProd.Recordset!NOMCORTO, "'")
        If concomillan > 0 Then
           XNOMCORTO = Mid(AdoProd.Recordset!NOMCORTO, 1, concomillan - 1) + "  " + Mid(AdoProd.Recordset!NOMCORTO, concomillan + 1, Len(AdoProd.Recordset!NOMCORTO))
        Else
           XNOMCORTO = AdoProd.Recordset!NOMCORTO
        End If
        'If Trim(Adoprod.Recordset!CONSEC) = "2603387" Then MsgBox "DFEFDSFSDF"
        clave = Str(Val(AdoProd.Recordset!CONSEC) - 1000000)
        Adodbf.Recordset.AddNew
        Adodbf.Recordset!CONSEC = Mid(Trim(clave), 1, 8) ' el campo no cabra en el registros
        Adodbf.Recordset!descripc = IIf(Not IsNull(AdoProd.Recordset!descripc), Mid(Trim(XDESCRIPC), 1, 50), " ")
        Adodbf.Recordset!NOMCORTO = IIf(Not IsNull(AdoProd.Recordset!NOMCORTO), Mid(XNOMCORTO, 1, 20), " ") 'el campo es demasiado paqueño para los datos
        Adodbf.Recordset!CONTENID = Val(IIf(Not IsNull(AdoProd.Recordset!CONTENID), AdoProd.Recordset!CONTENID, 0))
        Adodbf.Recordset!fletex = Val(IIf(Not IsNull(AdoProd.Recordset!fletex), AdoProd.Recordset!fletex, 0))
        Adodbf.Recordset!flesub = Val(IIf(Not IsNull(AdoProd.Recordset!flesub), AdoProd.Recordset!flesub, 0))
        Adodbf.Recordset!medida = IIf(Not IsNull(AdoProd.Recordset!medida), Mid(AdoProd.Recordset!medida, 1, 5), " ")
        Adodbf.Recordset!PAQUETES = Val(IIf(Not IsNull(AdoProd.Recordset!PAQUETES), AdoProd.Recordset!PAQUETES, 0))
        Adodbf.Recordset!peso = Val(IIf(Not IsNull(AdoProd.Recordset!peso), AdoProd.Recordset!peso, 0))
        Adodbf.Recordset!barraspza = Val(IIf(Not IsNull(AdoProd.Recordset!barraspza), AdoProd.Recordset!barraspza, 0))
        Adodbf.Recordset!barrascaja = Val(IIf(Not IsNull(AdoProd.Recordset!barrascaja), AdoProd.Recordset!barrascaja, 0))
        'POR DEFAULT EL DEPTO ES 1
        
        Adodbf.Recordset!tasaieps = Val(IIf(Not IsNull(AdoProd.Recordset!tasaieps), AdoProd.Recordset!tasaieps, 1))
        Adodbf.Recordset!claprove = IIf(Not IsNull(AdoProd.Recordset!claprove), Trim(AdoProd.Recordset!claprove), " ")
        Adodbf.Recordset!proceden = Val(IIf(Not IsNull(AdoProd.Recordset!procedencia), AdoProd.Recordset!procedencia, 0))
        Adodbf.Recordset!fecact = AdoProd.Recordset!fechaact
        'SI SE PASAN LAS PROMOCIONES
        Adodbf.Recordset!tantos = Val(IIf(Not IsNull(AdoProd.Recordset!cajas), AdoProd.Recordset!cajas, 0))
        Adodbf.Recordset!entre = Val(IIf(Not IsNull(AdoProd.Recordset!encajas), AdoProd.Recordset!encajas, 0))
        Adodbf.Recordset!clafamil = IIf(IsNull(AdoProd.Recordset!clafamil), "", AdoProd.Recordset!clafamil)
        Adodbf.Recordset!tasaiva = Val(IIf(Not IsNull(AdoProd.Recordset!tasaiva), AdoProd.Recordset!tasaiva, 0))
        'EL VERDADERO COSTO DE LA CAJA CON PROMOCIONES ESTA EN PRECOSTO
        'PARTE DE PRECIOS Y COSTO, EN FORMA PRECIOS DEBE GRABAR EL COSTO POR PIEZA
        Adodbf.Recordset!COSTOPAQ = Val(IIf(Not IsNull(AdoProd.Recordset!COSTOPAQ), AdoProd.Recordset!COSTOPAQ, 0))
        Adodbf.Recordset!costocaj = Val(IIf(Not IsNull(AdoProd.Recordset!PRECOSTO), AdoProd.Recordset!PRECOSTO, 0))
        Adodbf.Recordset!PRELISTA = Val(IIf(Not IsNull(AdoProd.Recordset!costocaj), AdoProd.Recordset!costocaj, 0))
        'LOS DESCUENTOS DEBEN SER DE LA TABLA DESCUENTOS
        Adodbf.Recordset!descto01 = Val(IIf(Not IsNull(AdoProd.Recordset!decto1), AdoProd.Recordset!decto1, 0))
        Adodbf.Recordset!descto02 = Val(IIf(Not IsNull(AdoProd.Recordset!decto2), AdoProd.Recordset!decto2, 0))
        Adodbf.Recordset!descto03 = Val(IIf(Not IsNull(AdoProd.Recordset!decto3), AdoProd.Recordset!decto3, 0))
        Adodbf.Recordset!descto04 = Val(IIf(Not IsNull(AdoProd.Recordset!dectoOferta), AdoProd.Recordset!dectoOferta, 0))
        Adodbf.Recordset!descto05 = Val(IIf(Not IsNull(AdoProd.Recordset!decto5), AdoProd.Recordset!decto5, 0))
        Adodbf.Recordset!descto06 = Val(IIf(Not IsNull(AdoProd.Recordset!dectoFinanciero), AdoProd.Recordset!dectoFinanciero, 0))
        Adodbf.Recordset!descefec = Val(IIf(Not IsNull(AdoProd.Recordset!dectoefectivo), AdoProd.Recordset!dectoefectivo, 0))
        'LOS CARGOS TAMBIEN SE DEBEN JALAR DE LOS CARGOS
        Adodbf.Recordset!porcargo = Val(IIf(Not IsNull(AdoProd.Recordset!cargo1), AdoProd.Recordset!cargo1, 0))
        Adodbf.Recordset!otrosrec = Val(IIf(Not IsNull(AdoProd.Recordset!cargo_efectivo), AdoProd.Recordset!cargo_efectivo, 0))
        Adodbf.Recordset!iva = Val(IIf(Not IsNull(AdoProd.Recordset!iva), AdoProd.Recordset!iva, 0))
        Adodbf.Recordset!ieps = Val(IIf(Not IsNull(AdoProd.Recordset!ieps), AdoProd.Recordset!ieps, 0))
        Adodbf.Recordset!fletes = Val(IIf(Not IsNull(AdoProd.Recordset!flete_efectivo), AdoProd.Recordset!flete_efectivo, 0))
        Adodbf.Recordset!maniobras = Val(IIf(Not IsNull(AdoProd.Recordset!maniobras), AdoProd.Recordset!maniobras, 0))
        Adodbf.Recordset!USUARIO = Mid(AdoProd.Recordset!prusuario, 1, 8)
        
        'LOS PRECIOS SE JALAN DE PREPROD
        Adodbf.Recordset!prepaque = Val(IIf(Not IsNull(AdoProd.Recordset!precio1), AdoProd.Recordset!precio1, 0))
        Adodbf.Recordset!PRECAJA = Val(IIf(Not IsNull(AdoProd.Recordset!PRECIO2), AdoProd.Recordset!PRECIO2, 0))
        Adodbf.Recordset!prelib1 = Val(IIf(Not IsNull(AdoProd.Recordset!PRECIO3), AdoProd.Recordset!PRECIO3, 0))
        Adodbf.Recordset!prelib2 = Val(IIf(Not IsNull(AdoProd.Recordset!precio4), AdoProd.Recordset!precio4, 0))
        Adodbf.Recordset!prelib3 = Val(IIf(Not IsNull(AdoProd.Recordset!precio5), AdoProd.Recordset!precio5, 0))
        Adodbf.Recordset!prelib4 = Val(IIf(Not IsNull(AdoProd.Recordset!precio6), AdoProd.Recordset!precio6, 0))
        'Ofertas solo para Fuchitan
        'Adodbf.Recordset!prepaque = Val(IIf(Not IsNull(AdoProd.Recordset!precio1), AdoProd.Recordset!precio1 - (AdoProd.Recordset!precio1 * 0.01), 0))
        'Adodbf.Recordset!PRECAJA = Val(IIf(Not IsNull(AdoProd.Recordset!PRECIO2), AdoProd.Recordset!PRECIO2 - (AdoProd.Recordset!PRECIO2 * 0.01), 0))
        'Adodbf.Recordset!prelib1 = Val(IIf(Not IsNull(AdoProd.Recordset!PRECIO3), AdoProd.Recordset!PRECIO3 - (AdoProd.Recordset!PRECIO3 * 0.01), 0))
        'Adodbf.Recordset!prelib2 = Val(IIf(Not IsNull(AdoProd.Recordset!precio4), AdoProd.Recordset!precio4 - (AdoProd.Recordset!precio4 * 0.01), 0))
        'Adodbf.Recordset!prelib3 = Val(IIf(Not IsNull(AdoProd.Recordset!precio5), AdoProd.Recordset!precio5 - (AdoProd.Recordset!precio5 * 0.01), 0))
        'Adodbf.Recordset!prelib4 = Val(IIf(Not IsNull(AdoProd.Recordset!precio6), AdoProd.Recordset!precio6 - (AdoProd.Recordset!precio6 * 0.01), 0))

        'LAS ESCALAS SE JALAN DE MARGEN
        Adodbf.Recordset!gananpaq = Val(IIf(Not IsNull(AdoProd.Recordset!escala1), AdoProd.Recordset!escala1, 0))
        Adodbf.Recordset!ganancaj = Val(IIf(Not IsNull(AdoProd.Recordset!escala2), AdoProd.Recordset!escala2, 0))
        Adodbf.Recordset!gananlib1 = Val(IIf(Not IsNull(AdoProd.Recordset!escala3), AdoProd.Recordset!escala3, 0))
        Adodbf.Recordset!gananlib2 = Val(IIf(Not IsNull(AdoProd.Recordset!escala4), AdoProd.Recordset!escala4, 0))
        'SE MANDA EL STATUS DE ACTIVO
        If AdoProd.Recordset!activo = 0 Then
          'MsgBox "producto desactivado"
           Adodbf.Recordset!barraspza = 0
           tt = 1
        Else
           tt = 0
        End If
        Adodbf.Recordset!pedir = tt
        If AdoProd.Recordset!OFERTADO Then
           'se agregan a un archivo de ofertas linea por producto
           des = Mid(AdoProd.Recordset!descripc, 1, 50)
           con = Val(IIf(Not IsNull(AdoProd.Recordset!CONTENID), AdoProd.Recordset!CONTENID, 0))
           med = IIf(Not IsNull(AdoProd.Recordset!medida), AdoProd.Recordset!medida, " ")
           paq = Val(IIf(Not IsNull(AdoProd.Recordset!PAQUETES), AdoProd.Recordset!PAQUETES, 0))
           pre = Val(IIf(Not IsNull(AdoProd.Recordset!PRECIO2), AdoProd.Recordset!PRECIO2, 0))
           caj = Val(IIf(Not IsNull(AdoProd.Recordset!PRECIO2), AdoProd.Recordset!PRECIO2, 0))
           linea = Mid(des, 1, 50) & " " & Format(con, "###0.000") & " " & med & "  " & Format(paq, "###0.00") & "  " & Format(pre, "#,###,####.00") & " " & Format(caj, "#,###,####.00") & " "
           'Write #1, linea
           f1 = 1
        Else
           f1 = 0
        End If
        Adodbf.Recordset!oferta = f1
        'LA CLAVE ESPECIAL DEL PROVEEDOR
        Adodbf.Recordset!claveprov = IIf(Not IsNull(AdoProd.Recordset!clavedelprov), AdoProd.Recordset!clavedelprov, 0)
        Adodbf.Recordset.Update
        'If opcion = "CAMBIOS" Then
        '   cn.Execute "UPDATE tfproduc SET actualizado = 0 WHERE consec = '" & Trim(Adoprod.Recordset!CONSEC) & "'"
        'End If
        suma = suma + 1
        avance.Caption = suma
        avance.Refresh
        AdoProd.Recordset.MoveNext
    Wend
Else
    MsgBox "No existieron cambios, VERIFIQUE SUS CORRECIONES", vbInformation
    Exit Sub
End If
'AdoDbf.Connect = "dbase III;"
'AdoDbf.DatabaseName = rutacambios
'AdoDbf.RecordsetType = Table
'AdoDbf.RecordSource = "catprov"
'AdoDbf.Refresh
Adodbf.RecordSource = "SELECT * FROM catprov"
Adodbf.Refresh

AdoProv.Refresh
probar1.Min = 0
probar1.Max = AdoProv.Recordset.RecordCount
lbltransO.Visible = True
lbltransO.Refresh
probar1o.Visible = True
v = 0
If AdoProv.Recordset.RecordCount > 0 Then
    AdoProv.Recordset.MoveFirst
    While Not AdoProv.Recordset.EOF
        v = v + 1
        probar1.Value = v
        Adodbf.Recordset.AddNew
        Adodbf.Recordset!prove = IIf(Not IsNull(AdoProv.Recordset!prove), Mid(Trim(AdoProv.Recordset!prove), 1, 3), " ")
        Adodbf.Recordset!NOMPROVE = IIf(Not IsNull(Mid(AdoProv.Recordset!NOMPROVE, 1, 50)), Mid(AdoProv.Recordset!NOMPROVE, 1, 50), " ")
        Adodbf.Recordset!dirpro = IIf(Not IsNull(AdoProv.Recordset!dirpro), Trim(AdoProv.Recordset!dirpro), " ")
        Adodbf.Recordset!colpro = IIf(Not IsNull(AdoProv.Recordset!colpro), Trim(AdoProv.Recordset!colpro), " ")
        Adodbf.Recordset!delpro = IIf(Not IsNull(AdoProv.Recordset!delpro), Trim(AdoProv.Recordset!delpro), " ")
        Adodbf.Recordset!codpro = IIf(Not IsNull(AdoProv.Recordset!codpro), Trim(AdoProv.Recordset!codpro), " ")
        Adodbf.Recordset!ciupro = IIf(Not IsNull(AdoProv.Recordset!ciupro), Trim(AdoProv.Recordset!ciupro), " ")
        Adodbf.Recordset!LOCPRO = IIf(Not IsNull(AdoProv.Recordset!LOCPRO), Trim(AdoProv.Recordset!LOCPRO), " ")
        Adodbf.Recordset!telpro = IIf(Not IsNull(AdoProv.Recordset!telpro), Trim(AdoProv.Recordset!telpro), " ")
        Adodbf.Recordset!frecuencia = IIf(Not IsNull(AdoProv.Recordset!frecuencia), AdoProv.Recordset!frecuencia, 0)
        Adodbf.Recordset!Volumen = IIf(AdoProv.Recordset!Volumen, 1, 0)
        Adodbf.Recordset!comprador = IIf(Not IsNull(AdoProv.Recordset!comprador), AdoProv.Recordset!comprador, " ")
        Adodbf.Recordset!activo = IIf(AdoProv.Recordset!activo, 1, 0)
        Adodbf.Recordset!tipo = Trim(AdoProv.Recordset!tipo)
        Adodbf.Recordset.Update
        AdoProv.Recordset.MoveNext
    Wend
End If
probar1o.Visible = False
lbltransO.Visible = False
lbltransO.Refresh
Adodbf.Recordset.Close

'AdoDbf.Connect = "dbase III;"
'AdoDbf.DatabaseName = "P:\paso"
'AdoDbf.RecordsetType = Table
'AdoDbf.RecordSource = "productox"
'AdoDbf.Refresh
Adodbf.RecordSource = "SELECT * FROM productox"
Adodbf.Refresh

Adodbf.Recordset.Close
Set Adodbf.Recordset = Nothing
'VERIFICAR SI TIENE DATOS EL ARCHIVO CAMBIOS
If Nivel = "I" Then  'Si es compras internas
   fs.copyfile "P:\pasoInt\producto.dbf", "P:\pasoInt\cambios.dbf", True
Else
   fs.copyfile "P:\paso\producto.dbf", "P:\paso\cambios.dbf", True
End If
If opcion = "CAMBIOS" Then
    respsn = MsgBox(" A CONTINUACION SE GENERARA EL ARCHIVO DE CAMBIOS PARA TIENDAS, DESEA SOBREESCRIBIR", vbQuestion + vbYesNoCancel, "Utilerias")
    If respsn = vbYes Then
        If Nivel = "I" Then  'Si es compras internas
           fs.copyfile "P:\pasoInt\producto.dbf", "P:\pasoInt\cambios.dbf", True
        Else
           fs.copyfile "P:\paso\producto.dbf", "P:\paso\cambios.dbf", True
        End If
    Else
        MsgBox "Por favor respalde el archivo cambios y agregue manualmente, APARTIR DEL ARCHIVO PRODUCTO.DBF", vbInformation
    End If
End If
MsgBox "PROCESO DE GENERACION DE CAMBIOS CONCLUIDO !!!", vbInformation
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub actcatprev()
Dim CNMDB As ADODB.Connection
Dim RST As ADODB.Recordset
Dim Rutas(1 To 9) As String
Rutas(1) = "RF1"
Rutas(2) = "RF2"
Rutas(3) = "RF3"
Rutas(4) = "RF4"
Rutas(5) = "RF5"
Rutas(6) = "RF6"
Rutas(7) = "RL1"
Rutas(8) = "RL2"
Rutas(9) = "RL3"
  Set CNMDB = New ADODB.Connection
  cruta = "P:\Preventa\Handheld\"
  cmen = stbmensajes.Panels(1).Text
  For x = 1 To 9
    cRutArc = cruta & Rutas(x)
    stbmensajes.Panels(1).Text = "Exportando catalogo de clientes " & cRutArc
    stbmensajes.Refresh
    If Dir(cruta & Rutas(x) & ".MDB") <> "" And lstrutas.Selected(x - 1) = True Then
      If x = 5 Then
         respor = InputBox("Porcentaje a incrementar a la ruta foránea 5?", "Porcentaje", 0)
         If Not IsNumeric(respor) Then
            MsgBox "El porcentaje debe ser númerico"
            Exit Sub
         End If
         respor = respor / 100
      End If
      CNMDB.Open "DSN=PITICOMDB;DBQ=" & cRutArc & ";DefaultDir=" & cruta & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
      CNMDB.Execute "DELETE FROM clientes"
      Set RST = New ADODB.Recordset

      cadsql = "SELECT CCLAVE, max(CNOMBRE) cnombre, max(nombres) nombres, max(apepaterno) apepaterno, max(apematerno) apematerno, max(nomnegocio) negocio, max(crfc) crfc, max(cdireccion) cdireccion, max(ccolonia) ccolonia, max(cciudad) cciudad, max(ctelefono) ctelefono, (CASE ccredito WHEN 1 THEN 1 END) credito, max(ruta) Ruta  FROM CATCLIENTE " & _
              "WHERE ruta LIKE '" & Rutas(x) & "%' GROUP BY cclave, (CASE ccredito WHEN 1 THEN 1 END)"
      AdoProd.CursorType = adOpenStatic
      AdoProd.LockType = adLockReadOnly
      AdoProd.ConnectionString = cCadConex
      AdoProd.RecordSource = cadsql
      AdoProd.Refresh
      N = 0:  probar1o.Max = AdoProd.Recordset.RecordCount
     While Not AdoProd.Recordset.EOF
        N = N + 1: lbltransO.Caption = N: lbltransO.Refresh
        probar1o.Value = N
        negocio = AdoProd.Recordset!negocio
        If InStr(1, AdoProd.Recordset!negocio, "'") Then negocio = Replace(AdoProd.Recordset!negocio, "'", " ")
        CNMDB.Execute "INSERT INTO clientes (clave,nombre,rfc,calle,colonia,poblacion,telefono, apepaterno,apematerno,nombres,nomnegocio,credito,ruta) VALUES (" & AdoProd.Recordset!cclave & ",'" & Replace(AdoProd.Recordset!cNombre, "'", " ") & "','" & AdoProd.Recordset!crfc & "','" & AdoProd.Recordset!cdireccion & "','" & AdoProd.Recordset!ccolonia & "','" & AdoProd.Recordset!cciudad & "','" & AdoProd.Recordset!ctelefono & "','" & AdoProd.Recordset!apepaterno & "','" & AdoProd.Recordset!apematerno & "','" & AdoProd.Recordset!nombres & "','" & negocio & "'," & IIf(IsNull(AdoProd.Recordset!credito) Or AdoProd.Recordset!credito = 0, 0, 1) & ",'" & AdoProd.Recordset!ruta & "')"
        AdoProd.Recordset.MoveNext
     Wend
     AdoProd.Recordset.Close
     stbmensajes.Panels(1).Text = "Exportando catalogo de productos " & cRutArc
     stbmensajes.Refresh
     CNMDB.Execute "DELETE FROM producto"
     AdoProd.CursorType = adOpenStatic
     AdoProd.LockType = adLockReadOnly
     AdoProd.ConnectionString = cCadConex
     AdoProd.RecordSource = "SELECT consec, descripc, substring(descripc,1,50), ltrim(STR(PAQUETES)) + ' X ' + ltrim(STR(CONTENID,8,3))+ ' ' + MEDIDA as present ,precio1,precio2,precio3,precio4,precio5,precio6,incant,incantpza,precosto FROM tfproduc,preprod,inventario WHERE consec = inprod AND consec = preclave AND inprod = preclave and (incant > 0 OR incantpza > 0)"
     AdoProd.Refresh
     N = 0:  probar1o.Max = AdoProd.Recordset.RecordCount
     While Not AdoProd.Recordset.EOF
        N = N + 1: lbltransO.Caption = N: lbltransO.Refresh
        probar1o.Value = N
        If Rutas(x) = "RF5" Then  'LA RUTA 5 FORANEA LEJANA SE AUMENTA EL 2% a cigarro y 5% a lo demas
           nporcent = IIf(InStr(1, AdoProd.Recordset!descripc, "CIGARRO") > 0, 1, 1 + respor)
           CNMDB.Execute "INSERT INTO producto (consec,descripc,medida,precio1 ,precio2 ,precio3 ,precio4, precio5,precio6,exicaj,exipza,costocaj) VALUES (" & AdoProd.Recordset!CONSEC & ",'" & Replace(AdoProd.Recordset!descripc, "'", " ") & "','" & AdoProd.Recordset!Present & "'," & AdoProd.Recordset!precio1 * nporcent & "," & AdoProd.Recordset!PRECIO2 * nporcent & "," & AdoProd.Recordset!PRECIO3 * nporcent & "," & AdoProd.Recordset!precio4 * nporcent & "," & IIf(IsNull(AdoProd.Recordset!precio5), 0, AdoProd.Recordset!precio5 * nporcent) & "," & IIf(IsNull(AdoProd.Recordset!precio6), 0, AdoProd.Recordset!precio6 * nporcent) & "," & AdoProd.Recordset!InCant & "," & AdoProd.Recordset!InCantPza & "," & AdoProd.Recordset!PRECOSTO & ")"
        Else
           'MsgBox "INSERT INTO producto (consec,descripc,medida,precio1,precio2,precio3,precio4,precio5,precio6,exicaj,exipza,costocaj) VALUES (" & Adoprod.Recordset!CONSEC & ",'" & Replace(Adoprod.Recordset!descripc, "'", " ") & "','" & Adoprod.Recordset!Present & "'," & Adoprod.Recordset!precio1 & "," & Adoprod.Recordset!PRECIO2 & "," & Adoprod.Recordset!PRECIO3 & "," & Adoprod.Recordset!precio4 & "," & IIf(IsNull(Adoprod.Recordset!precio5), 0, Adoprod.Recordset!precio5) & "," & IIf(IsNull(Adoprod.Recordset!precio6), 0, Adoprod.Recordset!precio6) & "," & Adoprod.Recordset!InCant & "," & Adoprod.Recordset!InCantPza & "," & Adoprod.Recordset!PRECOSTO & ")"
           CNMDB.Execute "INSERT INTO producto (consec,descripc,medida,precio1,precio2,precio3,precio4,precio5,precio6,exicaj,exipza,costocaj) VALUES (" & AdoProd.Recordset!CONSEC & ",'" & Replace(AdoProd.Recordset!descripc, "'", " ") & "','" & AdoProd.Recordset!Present & "'," & AdoProd.Recordset!precio1 & "," & AdoProd.Recordset!PRECIO2 & "," & AdoProd.Recordset!PRECIO3 & "," & AdoProd.Recordset!precio4 & "," & IIf(IsNull(AdoProd.Recordset!precio5), 0, AdoProd.Recordset!precio5) & "," & IIf(IsNull(AdoProd.Recordset!precio6), 0, AdoProd.Recordset!precio6) & "," & AdoProd.Recordset!InCant & "," & AdoProd.Recordset!InCantPza & "," & AdoProd.Recordset!PRECOSTO & ")"
        End If
        AdoProd.Recordset.MoveNext
     Wend
     CNMDB.Close
     stbmensajes.Panels(1).Text = "Compactando base de datos " & Rutas(x)
     stbmensajes.Refresh
     sNameDest = cRutArc + ".BAK"
     If Dir(sNameDest) <> "" Then
        Kill sNameDest
     End If
     sNameOrig = cRutArc + ".MDB"
     vEscenario = dbLangSpanish    ' en el supuesto de BD en Español
     'vContra = ";pwd=PORTATIL"   ' Atencion: no olvides el punto y coma
     DBEngine.CompactDatabase sNameOrig, sNameDest, vEscenario
     Dim fs As Object
     Set fs = CreateObject("Scripting.FileSystemObject")
     fs.copyfile sNameDest, sNameOrig, True
     Kill sNameDest
   End If
  Next
  stbmensajes.Panels(1) = cmen
  stbmensajes.Refresh
  probar1o.Visible = False
  MsgBox "LA EXPORTACION DE CATALOGOS DE PREVENTISTAS SE REALIZO CORRECTAMENTE", vbInformation
End Sub

Private Sub exportatodo()
'SE DEBEN PONER LOS DESCUENTOS VIGENTES, IVA VIGENTE Y NO QUE SE APLICAN A LOS PEDIDOS
AdoProd.CursorType = adOpenKeyset
AdoProd.LockType = adLockOptimistic
AdoProd.CommandType = adCmdText
AdoProd.ConnectionString = strconnect
RESP = MsgBox("Desea Generar Base de Datos Solo con productos Activos [NO = Incluir Inactivos] ", vbYesNo + vbQuestion)
If RESP = vbYes Then
   CADENA = "select * from tfproduc,descuentos,preprod,cargos,margen  where consec = deprod and consec=margen.producto and consec = preclave  and caprod = consec  and activo = 1 " & compInt & " ORDER BY consec"
Else
   CADENA = "select * from tfproduc,descuentos,preprod,cargos,margen  where consec = deprod and consec=margen.producto and consec = preclave  and caprod = consec  " & compInt
End If
AdoProd.RecordSource = CADENA
AdoProd.Refresh
Exporta ("TODO")
End Sub

Private Sub txtinicio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtfinal.SetFocus
ElseIf KeyAscii = 27 Then
   fraperi.Visible = False
End If

End Sub

Private Sub verpr_Click()
precycosto
End Sub

Private Sub vtamay_Click()
  frmPrincipal.Show
End Sub

Sub Importa(clave As String)
'On Error Resume Next
'LA IMPORTACION SIEMPRE DEBE SER TOTAL
'respsn = MsgBox("El archivo cambios.dbf , debe estar ubicado en P:\PITICO\PITICO10" & Chr(13) & "Deseas continuar... ? ", vbQuestion + vbYesNo, "Utilerias")
If respsn = vbCancel Then Exit Sub
Dim v As Integer
Dim rs As ADODB.Recordset
'IMPOFI
Set rs = New ADODB.Recordset
AdoProd.CursorType = adOpenKeyset
AdoProd.ConnectionString = "DSN=PITICOMDB;DBQ=P:\PITICO\PITICO10\PITICO10.mdb;DefaultDir=P:\PITICO\PITICO10;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"
'AdoProd.ConnectionString = "DSN=PITICOMDB;DBQ=c:\PITICO10.mdb;DefaultDir=c:\;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=PORTATIL;UID=admin;"
ctfp = IIf(SoloAct, "tfproduc", "tfproduc")
AdoProd.RecordSource = "SELECT * FROM " & ctfp & " WHERE CONSEC = '" & clave & "'"
AdoProd.Refresh
Dim cnsql As ADODB.Connection
Set cnsql = New ADODB.Connection
cnsql.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PITICO;Data Source=" & SERVIDOR
cnsql.Open
While Not AdoProd.Recordset.EOF
   rs.Open "SELECT * FROM tfproduc WHERE consec = '" & Trim(AdoProd.Recordset!CONSEC) & "'", cnsql, adOpenForwardOnly, adLockOptimistic, admcdtext
   If rs.BOF And rs.EOF Then
      If MsgBox("DESEAS DAR DE ALTA EL PRODUCTO " & AdoProd.Recordset!descripc, vbYesNoCancel + vbQuestion) = vbYes Or vbCancel Then
         clave = AdoProd.Recordset!CONSEC
         activo = IIf(vbCancel, 0, 1)
         cadsql = "INSERT INTO tfproduc(consec,claprove,descripc,nomcorto,paquetes,contenid,medida,barraspza,linea,costocaj,activo,fechaintro,igualofi) VALUES ('" & _
            clave & "','" & AdoProd.Recordset!claprove & "','" & Replace(AdoProd.Recordset!descripc, "'", " ") & "','" & Replace(AdoProd.Recordset!descripc, "'", " ") & "'," & _
            AdoProd.Recordset!PAQUETES & "," & AdoProd.Recordset!CONTENID & ",'" & AdoProd.Recordset!medida & "'," & AdoProd.Recordset!barraspza & ",'" & AdoProd.Recordset!clafamil & "'," & "0" & "," & activo & ",'" & date & "',1)"
         cnsql.Execute cadsql
         rs.Close
         rs.Open "SELECT * FROM PREPROD WHERE PRECLAVE = '" & Trim(AdoProd.Recordset!CONSEC) & "'", cnsql, adOpenForwardOnly, adLockOptimistic, admcdtext
         If rs.BOF And rs.EOF Then
            cnsql.Execute "INSERT INTO preprod(preclave) VALUES('" & clave & "')"
         End If
         rs.Close
         rs.Open "SELECT * FROM margen WHERE producto = '" & Trim(AdoProd.Recordset!CONSEC) & "'", cnsql, adOpenForwardOnly, adLockOptimistic, admcdtext
         If rs.BOF And rs.EOF Then
            cnsql.Execute "INSERT INTO margen(producto) VALUES('" & clave & "')"
         End If
      End If
   End If
   rs.Close
   AdoProd.Recordset.MoveNext
Wend
If MsgBox("DESEAS ACTUALIZAR CATALOGO DE PROVEEDORES", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
   AdoProd.Recordset.Close
   AdoProd.RecordSource = "SELECT * FROM CATPROV"
   AdoProd.Refresh
   While Not AdoProd.Recordset.EOF
     rs.Open "SELECT * FROM catprov WHERE PROVE = '" & AdoProd.Recordset!prove & "'", cnsql, adOpenDynamic, adLockOptimistic, admcdtext
     If rs.BOF And rs.EOF Then
        'cadsql = "INSERT INTO catprov(prove,nomprove,dirpro,colpro,delpro,codpro,ciupro,locpro,telpro,frecuencia,tipo,rfc,visita,procedencia,razon) VALUES ('" & Adoprod.Recordset!prove & "','" & Adoprod.Recordset!Nomprove & "','" & Adoprod.Recordset!dirpro & "','" & Adoprod.Recordset!colpro & "','" & Adoprod.Recordset!delpro & "','" & Adoprod.Recordset!codpro & "','" & Adoprod.Recordset!ciupro & "','" & IIf(IsNull(Adoprod.Recordset!LOCPRO), "X", Adoprod.Recordset!LOCPRO) & "','" & Adoprod.Recordset!telpro & "'," & IIf(IsNull(Adoprod.Recordset!frecuencia), 0, Adoprod.Recordset!frecuencia) & ",'" & Adoprod.Recordset!tipo _
                 & "','" & Adoprod.Recordset!rfc & "'," & IIf(IsNull(Adoprod.Recordset!visita), 0, Adoprod.Recordset!visita) & "," & IIf(IsNull(Adoprod.Recordset!procedencia), 0, Adoprod.Recordset!procedencia) & ",'" & Adoprod.Recordset!RAZON & "')"
        If Not IsNull(AdoProd.Recordset!NOMPROVE) Then
           cadsql = "INSERT INTO catprov(prove,nomprove,dirpro,colpro,delpro,codpro,ciupro,locpro,telpro) VALUES ('" & AdoProd.Recordset!prove & "','" & Replace(AdoProd.Recordset!NOMPROVE, "'", " ") & "','" & IIf(IsNull(AdoProd.Recordset!dirpro), ".", AdoProd.Recordset!dirpro) & "','" & AdoProd.Recordset!colpro & "','" & AdoProd.Recordset!delpro & "','" & AdoProd.Recordset!codpro & "','" & AdoProd.Recordset!ciupro & "','" & IIf(IsNull(AdoProd.Recordset!LOCPRO), "X", AdoProd.Recordset!LOCPRO) & "','" & AdoProd.Recordset!telpro & "')"
           cnsql.Execute cadsql
        End If
     End If
     AdoProd.Recordset.MoveNext
     rs.Close
   Wend
End If
MsgBox "LA ACTUALIZACION SE REALIZO CORRECTAMENTE", vbInformation, "Actualización"
Exit Sub
DaoProdDbf.Connect = "dbase III;"
'DaoProdDbf.DatabaseName = "P:\paso"
DaoProdDbf.DatabaseName = "P:\PASO"
DaoProdDbf.RecordsetType = Table
DaoProdDbf.RecordSource = "catprov"
DaoProdDbf.Refresh
clave = ""
If DaoProdDbf.Recordset.RecordCount > 0 Then
    probar1.Min = 0
    probar1.Max = DaoProdDbf.Recordset.RecordCount
    Lbltrans.Visible = True
    Lbltrans.Refresh
    probar1.Visible = True
    v = 0
    DaoProdDbf.Recordset.MoveFirst
    AdoProv.Refresh
    While Not DaoProdDbf.Recordset.EOF
      v = v + 1
      probar1.Value = v
      clave = DaoProdDbf.Recordset!prove
      concomilla = InStr(1, DaoProdDbf.Recordset!NOMPROVE, "'")
      If concomilla > 0 Then
         xnomprove = Mid(DaoProdDbf.Recordset!NOMPROVE, 1, concomilla - 1) + "  " + Mid(DaoProdDbf.Recordset!NOMPROVE, concomilla + 1, Len(DaoProdDbf.Recordset!NOMPROVE))
      Else
         xnomprove = DaoProdDbf.Recordset!NOMPROVE
      End If
      xnomprove = Mid(xnomprove, 1, 150)
      AdoProv.Recordset.MoveFirst
      AdoProv.Recordset.Find "prove = '" & Trim(clave) & "'"
      If AdoProv.Recordset.EOF Then
        CADENA = "INSERT INTO CATPROV (prove,nomprove,tipo,ACTIVO,BACKORDER) values ('" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!prove), DaoProdDbf.Recordset!prove, 0) & "','" & IIf(Not IsNull(DaoProdDbf.Recordset!NOMPROVE), xnomprove, 0) & "'," & _
        "'I',1,1)"
        
        cn.Execute CADENA
      Else
        CADENA = "UPDATE catprov SET nomprove = '" & IIf(Not IsNull(Mid(DaoProdDbf.Recordset!NOMPROVE, 1, 150)), xnomprove, 0) & "', dirpro = '" & _
        IIf(Not IsNull(Mid(DaoProdDbf.Recordset!dirpro, 1, 50)), Mid(DaoProdDbf.Recordset!dirpro, 1, 50), 0) & "', colpro = '" & IIf(Not IsNull(DaoProdDbf.Recordset!colpro), DaoProdDbf.Recordset!colpro, 0) & "', delpro =  '" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!delpro), DaoProdDbf.Recordset!delpro, 0) & "', codpro =  '" & IIf(Not IsNull(DaoProdDbf.Recordset!codpro), DaoProdDbf.Recordset!codpro, 0) & "', ciupro = '" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!ciupro), DaoProdDbf.Recordset!ciupro, 0) & "', locpro = '" & IIf(Not IsNull(DaoProdDbf.Recordset!LOCPRO), DaoProdDbf.Recordset!LOCPRO, 0) & "', telpro = '" & _
        IIf(Not IsNull(DaoProdDbf.Recordset!telpro), DaoProdDbf.Recordset!telpro, 0) & "' where prove = '" & Trim(clave) & "'"
        'MsgBox cadena
        cn.Execute CADENA
      End If
      DaoProdDbf.Recordset.MoveNext
     Wend
End If
probar1.Visible = False
Lbltrans.Visible = False
Lbltrans.Refresh

'DaoProdDbf.Connect = "dbase III;"
'DaoProdDbf.DatabaseName = "P:\paso"
''DaoProdDbf.RecordsetType = Table
'DaoProdDbf.RecordSource = "productox"
'DaoProdDbf.Refresh

MsgBox "Es necesario salir del programa , para actualizar los cambios realizados...", vbInformation
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub GencatCaj()
Dim rs As ADODB.Recordset
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
origen = "\\" & SERVIDOR & "\disco-c\programas\escajera.dbf"
destino = "\\" & SERVIDOR & "\disco-c\vliventa\usuarios.dbf"
fs.copyfile origen, destino
Daocambio.Connect = "dbase III;"
Daocambio.DatabaseName = "\\" & SERVIDOR & "\disco-c\vliventa"
Daocambio.RecordsetType = vbRSTypeDynaset
Daocambio.RecordSource = "usuarios"
Daocambio.Refresh
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM usuarios WHERE level1 = 'J'", cn, adOpenForwardOnly, adLockOptimistic, admdtext
While Not rs.EOF
   Daocambio.Recordset.AddNew
   Daocambio.Recordset!login = Trim(rs!login)
   Daocambio.Recordset!Name = Trim(rs!Name)
   Daocambio.Recordset!clave = rs!clave
   Daocambio.Recordset!sucursal = rs!sucursal
   Daocambio.Recordset.Update
   rs.MoveNext
Wend
rs.Close
Daocambio.RecordSource = "catprov"
Daocambio.Refresh
Set rs = Nothing
Daocambio.Recordset.Close
Set Daocambio.Recordset = Nothing
MsgBox "LA EXPORTACION DEL CATALOGO DE CAJERAS/CAJEROS SE REALIZO CORRECTAMENTE", vbInformation
End Sub


Private Sub rptcorteInv()
Dia = "dia" & Day(fecini.Value)
'Igualo el inventario del dia de inventario con el inventario del dia
cn.Execute "UPDATE inventario SET ininidia = 0, ininidiap = 0, entpedp = 0, enttra = 0, ajustes = 0, salvta = 0, saltra = 0"
cn.Execute "UPDATE inventario SET inIniDia = " & Dia & " FROM invcorte WHERE producto = inprod and mes = " & Month(date)
'Entradas a través de Pedidos por proveedor
cn.Execute "DELETE FROM invtemp"
cn.Execute "INSERT INTO invtemp SELECT dg_producto, sum(dg_cantreal) + SUM(dg_promocionr) FROM Pedprove P, detalleglobal d WHERE p.pp_pedido = d.dg_pedido AND p.pp_fecrecibe >='" & fecini.Value & "' and  p.pp_fecrecibe <= '" & DateAdd("d", 1, fecfin.Value) & "' GROUP BY dg_producto"
cn.Execute "UPDATE inventario SET entpedp = t.Totmov FROM invtemp t WHERE producto = inprod"
'Entradas a través de traslados
cn.Execute "DELETE FROM invtemp"
cn.Execute "INSERT INTO invtemp SELECT dt_producto, sum(dt_cantidad) FROM traslados t, detalletraslado d WHERE t.t_clave = d.dt_clave AND t.T_fecha >='" & fecini.Value & "' and  t.t_fecha <= '" & DateAdd("d", 1, fecfin.Value) & "' and t.t_enviado = 1 and t_entrada = 1 and t_motivocancela is null GROUP BY dt_producto"
cn.Execute "UPDATE inventario SET enttra = t.Totmov FROM invtemp t WHERE producto = inprod"
'Salidas/Entradas a través de Ajustes
cn.Execute "DELETE FROM invtemp"
cn.Execute "INSERT INTO invtemp SELECT da_producto, sum(da_cantidad) FROM Ajustes A, detalleAjustes d WHERE a.a_clave = d.da_clave AND a.a_fecha >='" & fecini.Value & "' and  a.a_fecha <= '" & DateAdd("d", 1, fecfin.Value) & "' GROUP BY da_producto"
cn.Execute "UPDATE inventario SET ajustes = t.Totmov FROM invtemp t WHERE producto = inprod"
'Salidas a través de ventas
cn.Execute "DELETE FROM invtemp"
cn.Execute "INSERT INTO invtemp SELECT cl_producto, sum(cantidad) FROM ventas v, ventas_det d WHERE v.noventa = d.noventa AND v.fecha >='" & fecini.Value & "' and  v.fecha <= '" & DateAdd("d", 1, fecfin.Value) & "' and d.cancelado = 0 GROUP BY cl_producto"
cn.Execute "UPDATE inventario SET salvta = t.Totmov FROM invtemp t WHERE producto = inprod"
'Salidas a través de traslados
cn.Execute "DELETE FROM invtemp"
cn.Execute "INSERT INTO invtemp SELECT dt_producto, sum(dt_cantidad) FROM traslados t, detalletraslado d WHERE t.t_clave = d.dt_clave AND t.T_fecha >='" & fecini.Value & "' and  t.t_fecha <= '" & DateAdd("d", 1, fecfin.Value) & "' and t.t_enviado = 1 and t_entrada = 0 and t_motivocancela is null GROUP BY dt_producto"
cn.Execute "UPDATE inventario SET saltra = t.Totmov FROM invtemp t WHERE producto = inprod"
CR1.WindowTitle = "Rotacion de productos"
CR1.ReportFileName = App.Path & "\INVCORTE.RPT"
CR1.Formulas(0) = "ENCABEZADO = 'ROTACION DE PRODUCTOS DEL " & fecini.Value & " AL " & fecfin.Value & "'"
CR1.SQLQuery = ""
If MsgBox("DESEAS VER EL REPORTE SOLO CON AQUELLOS PRODUCTOS CON MOVIMIENTOS", vbYesNo + vbQuestion) = vbYes Then
   'cr1.Formulas(0) = "FORMSELEC = {@INVFIN} <> {INVENTARIO.incant} AND {@INVFIN} > 0 "
   CR1.Formulas(0) = "FORMSELEC = ({@TOTENT} > 0 OR {@TOTSAL} > 0) "
Else
   CR1.Formulas(0) = "FORMSELEC = {INVENTARIO.incant} > 0  OR ({@TOTENT} > 0 OR {@TOTSAL} > 0)"
End If
CR1.Formulas(1) = "ENCABEZADO = 'REPORTE DE MOVIMIENTOS DEL " & fecini.Value & " AL " & fecfin.Value & "'"
CR1.Connect = cn.ConnectionString
CR1.Action = 1
End Sub

Private Sub importadesp()
Dim dtas(1 To 2) As String
dtas(1) = "DESCOS"
dtas(2) = "DESCEN"

'SE HACE UN CICLO PARA IMPORTAR LOS INVENTARIOS DE LAS TIENDAS
ANIMA
For i = 1 To 2
  If Dir("P:\BUZON\" & dtas(i) & ".TXT") <> "" Then
     Open "P:\BUZON\" & dtas(i) & ".TXT" For Input As #1
     cArch = dtas(i)
     sucu = Pedsuc(dtas(i))
     stbmensajes.Panels(1).Text = "Importando despensas " & cArch & " ..."
     stbmensajes.Refresh
     lblProd.Visible = True
     nreg = 1
     Line Input #1, CAD
     dfecinI = Mid(CAD, Len(CAD) - 7, 8)
     While Not EOF(1)
       Line Input #1, CAD
       nreg = nreg + 1
     Wend
     Close #1
     dfecfin = Mid(CAD, Len(CAD) - 7, 8)
     probar1.Min = 0
     probar1.Max = nreg
     nreg = 0
     'SE INICIALIZA EL INVENTARIO DE LA TIENDA CORRESPONDIENTE
     'If dtas(i) = "DESCEN" Then
     '   SUC = "SUCURSAL = '23' "
     'ElseIf dtas(i) = "DESCOS" Then
     '   CONDSER = "SUCURSAL= '24'"
     'End If
     cn.Execute "DELETE FROM DESPSUC WHERE sucursal = " & sucu
     Open "P:\BUZON\" & dtas(i) & ".TXT" For Input As #1
     While Not EOF(1)
        Line Input #1, CAD
        nreg = nreg + 1
        probar1.Value = nreg
        lblProd.Caption = "Registros procesados: " & Str(nreg): lblProd.Refresh
         pos1 = InStr(CAD, "|")
        Folio = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
         pos1 = InStr(CAD, "|")
        clave = Mid(CAD, 1, pos1 - 1)
         CAD = Mid(CAD, pos1 + 1, Len(CAD))
        cantidad = CAD
        cn.Execute "INSERT INTO DESPSUC(folio,clave,cantidad,sucursal) VALUES (" & Folio & ",'" & clave & "'," & cantidad & ",'" & sucu & "')"
      Wend
      Close #1
  End If
Next
MsgBox "LA IMPORTACION SE REALIZO CORRECTAMNETE", vbInformation, "Utilerías"
Me.ani1.Close
Me.ani1.Visible = False
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim DatosRecibidos As String
  Winsock1.GetData DatosRecibidos, vbString
  MsgBox "Dirección pública asignada a esta sesión en Internet" & Chr(13) & Chr(13) & Space(20) & DatosRecibidos, vbInformation, "Ip Pública"
End Sub

