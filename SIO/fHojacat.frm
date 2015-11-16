VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fhojacat 
   Caption         =   "Hoja de Catalogo"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "fHojacat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FRACONTRA 
      BackColor       =   &H00808000&
      Caption         =   "Escriba la contraseña..."
      Enabled         =   0   'False
      Height          =   1335
      Left            =   2880
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtcontra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdconfirma 
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6855
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   11655
      Begin MSDataGridLib.DataGrid DGrprod 
         Bindings        =   "fHojacat.frx":0442
         Height          =   5055
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1.5
         RowHeight       =   15
         TabAction       =   1
         RowDividerStyle =   6
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
         ColumnCount     =   28
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "barraspza"
            Caption         =   "Cod. Barras"
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
         BeginProperty Column03 
            DataField       =   "descripc"
            Caption         =   "Nombre del Producto"
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
            DataField       =   "paquetes"
            Caption         =   "Pz X Caja"
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
            DataField       =   "contenid"
            Caption         =   "Pres."
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
         BeginProperty Column07 
            DataField       =   "costocaj"
            Caption         =   "Precio Lista"
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
            DataField       =   "precosto"
            Caption         =   "Precio Costo"
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
            DataField       =   "precio_v"
            Caption         =   "Precio Venta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "incant"
            Caption         =   "Inventario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "decto1"
            Caption         =   "Descto.1"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "decto2"
            Caption         =   "Descto. 2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "decto3"
            Caption         =   "Descto. 3"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "efectivo"
            Caption         =   "Efectivo"
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
         BeginProperty Column15 
            DataField       =   "decto4"
            Caption         =   "Descto. 4"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "financiero"
            Caption         =   "Descto. Finan."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "cajas"
            Caption         =   "Promo. Cajas"
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
         BeginProperty Column18 
            DataField       =   "encajas"
            Caption         =   "En Cajas"
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
            DataField       =   "plazopago"
            Caption         =   "Dias Plazo "
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
         BeginProperty Column20 
            DataField       =   "cargo3"
            Caption         =   "IVA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column21 
            DataField       =   "cargo4"
            Caption         =   "IEPS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column22 
            DataField       =   "flete"
            Caption         =   "Flete"
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
         BeginProperty Column23 
            DataField       =   "maniobras"
            Caption         =   "Maniobras"
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
         BeginProperty Column24 
            DataField       =   "CARGO1"
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
         BeginProperty Column25 
            DataField       =   "CARGO2"
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
         BeginProperty Column26 
            DataField       =   "CARGO5"
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
         BeginProperty Column27 
            DataField       =   "clavedelprov"
            Caption         =   "Clav. Esp."
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
         SplitCount      =   2
         BeginProperty Split0 
            ScrollBars      =   1
            Size            =   496
            BeginProperty Column00 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   3270.047
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnAllowSizing=   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column12 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column13 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column14 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column15 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column16 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column17 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column18 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column19 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column20 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column21 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column22 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column23 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column24 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column25 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column26 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column27 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
         BeginProperty Split1 
            ScrollBars      =   3
            RecordSelectors =   0   'False
            ScrollGroup     =   0
            Size            =   246
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   3270.047
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column12 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column13 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column14 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column15 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column16 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column17 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column18 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column19 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column20 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column21 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column22 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column23 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column24 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column25 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column26 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column27 
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         DataField       =   "activo"
         DataSource      =   "Adoprod"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   29
         Top             =   1320
         Width           =   255
      End
      Begin VB.ComboBox Cmbprov 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox TxtProv 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox Cmbfamilia 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox Txtfamilia 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Txtdepto 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   3855
      End
      Begin VB.CheckBox Chkprocede 
         Alignment       =   1  'Right Justify
         Caption         =   "    Local"
         Enabled         =   0   'False
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
         Left            =   9720
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.PictureBox CR1 
         Height          =   480
         Left            =   7320
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   31
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lbletiquetas 
         Alignment       =   2  'Center
         Caption         =   "Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label producto 
         Alignment       =   2  'Center
         Caption         =   "."
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
         Left            =   960
         TabIndex        =   28
         Top             =   1200
         Width           =   6615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Familias     :"
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
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento:"
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
         Index           =   2
         Left            =   7560
         TabIndex        =   17
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Proveedor:"
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
         Index           =   3
         Left            =   8040
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Plazo de Pago"
      Height          =   975
      Left            =   9120
      TabIndex        =   6
      Top             =   6960
      Width           =   1335
      Begin VB.TextBox Txtplazo 
         Height          =   375
         Left            =   120
         MaxLength       =   3
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc Adousuario 
      Height          =   330
      Left            =   2280
      Top             =   8160
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
      Caption         =   "usuario"
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
   Begin MSAdodcLib.Adodc Adoprov 
      Height          =   330
      Left            =   7800
      Top             =   8160
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "proveed"
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
   Begin MSAdodcLib.Adodc Adoprod 
      Height          =   330
      Left            =   5040
      Top             =   8160
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
      Caption         =   "productos"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton CmdPREVER 
         Caption         =   "&Preeliminar"
         Height          =   855
         Left            =   7680
         Picture         =   "fHojacat.frx":0458
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdprecios 
         Caption         =   "Precios"
         Enabled         =   0   'False
         Height          =   855
         Left            =   8640
         Picture         =   "fHojacat.frx":098A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdpor 
         Caption         =   "%%%"
         Enabled         =   0   'False
         Height          =   855
         Left            =   9600
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Inventario"
         Enabled         =   0   'False
         Height          =   855
         Left            =   4440
         Picture         =   "fHojacat.frx":0DCC
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Modificar datos del Producto"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cond. de Compra"
         Enabled         =   0   'False
         Height          =   855
         Left            =   3360
         Picture         =   "fHojacat.frx":120E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Modificar datos del Producto"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "P&roducto"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2280
         Picture         =   "fHojacat.frx":1650
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Modificar datos del Producto"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Proveedor"
         Height          =   855
         Left            =   1200
         Picture         =   "fHojacat.frx":195A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modificar datos del Proveedor"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Precio&s"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "fHojacat.frx":1D9C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Modificar Precios por Producto"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   10560
         Picture         =   "fHojacat.frx":20A6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir del Modulo"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "fhojacat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset


Private Sub Adoprod_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo Error:
 producto.Caption = Adoprod.Recordset!descripc & Chr(13) & Adoprod.Recordset!PAQUETES & "  x " & Adoprod.Recordset!CONTENID & " " & Adoprod.Recordset!medida
 producto.Refresh
Exit Sub
Error:
  MsgBox "No existen producto de este proveedor..."
End Sub

Private Sub Cmbfamilia_Click()
On Error GoTo Error:
Dim N As Integer
   If Trim(Cmbfamilia.Text) <> "" Then
      N = InStr(1, Cmbfamilia.List(Cmbfamilia.ListIndex), "[")
      cvefamilia = Mid(Cmbfamilia.List(Cmbfamilia.ListIndex), N + 1, Len(Cmbfamilia.List(Cmbfamilia.ListIndex)) - N - 1)
      Txtfamilia.Text = cvefamilia
      N = InStr(1, Cmbfamilia.List(i), "(")
       
      sdepto = Mid(Cmbfamilia.List(i), N + 1, Len(Cmbfamilia.List(i)) - N - 1)
      Txtdepto.Text = sdepto
       
      ' ESTA ES LA FAMILIA, PERO SE DEBE BUSCAR LA RELACION ULTIMA ES POR LINEA
      ' DEPTO-> FAMILIA-> LINEA-> PRODUCTO
      ' ENTONCES SE DEBE BUSCAR LAS LINEAS QUE TIENE ESTE FAMILIA DE ESTE PROVEEDOR
      Adoprod.CursorLocation = adUseServer
      Adoprod.LockType = adLockOptimistic
      Adoprod.CursorType = adOpenKeyset
      activo = IIf(SoloAct, "AND TFPRODUC.ACTIVO = 1", "")
      Adoprod.RecordSource = "SELECT consec,costototal,claprove,paquetes,barraspza,clafamil,linea,descripc,contenid,medida,decto1,decto2,decto3,decto4,financiero,efectivo,descprod.cajas,descprod.encajas,cargo1,cargo2,cargo3,cargo4,cargo5,flete,maniobras,tfproduc.activo," & _
                             "descprod.plazopago,costocaj,precosto,precio1,precio1 * paquetes as precio_v,((( precio1 * paquetes )- costototal)/( precio1 * paquetes ))*100 as margen " & _
                             "FROM tfproduc,descprod,catprov,preprod WHERE consec=preclave " & activo & " AND consec=producto and claprove = prove and  claprove = '" & Trim(TxtProv.Text) & "' AND CLAFAMIL = '" & Trim(Txtfamilia.Text) & "' order by descripc,contenid"
      Adoprod.ConnectionString = strconnect
   
      Adoprod.Refresh
      If Adoprod.Recordset.RecordCount > 0 Then
        Command2.Enabled = True
      Else
        Command2.Enabled = False
      End If
        
   Else
       MsgBox "Seleccione una familia "
       Cmbfamilia.SetFocus
   End If

Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Cmbprov_Click()
On Error GoTo Error:
Dim RSTEMP As ADODB.Recordset
Dim clave1 As String
Dim N As Integer
Dim activo As String
   Txtdepto.Text = ""
   If Trim(cmbProv.Text) <> "" Then
      N = InStr(1, cmbProv.List(cmbProv.ListIndex), "[")
      cveprov = Mid(cmbProv.List(cmbProv.ListIndex), N + 1, Len(cmbProv.List(cmbProv.ListIndex)) - N - 1)
      TxtProv.Text = cveprov
      Set RSTEMP = New ADODB.Recordset
      activo = IIf(SoloAct, " AND activo = 1", "")
      RSTEMP.Open "SELECT * FROM catprov WHERE prove = '" & Trim(cveprov) & "'" & activo, cn, adOpenKeyset, adLockOptimistic, adCmdText
      Chkprocede.Value = RSTEMP!procedencia
      clave1 = IIf(Not IsNull(RSTEMP!comprador), RSTEMP!comprador, "000")
      RSTEMP.Close
    
      llenaprod (cveprov)
      DGrprod.SetFocus
    Else
       MsgBox "Seleccione un proveedor "
       cmbProv.SetFocus
    End If

Exit Sub
Error:
MsgBox Err.Description
End Sub


Private Sub cmdconfirma_Click()
If txtContra = "CAMBIO01" Then
   Call procesaescala
   FRACONTRA.Enabled = False
   FRACONTRA.Visible = False
End If
txtContra.Text = ""
FRACONTRA.Enabled = False
FRACONTRA.Visible = False
End Sub

Private Sub cmdpor_Click()
Me.FRACONTRA.Enabled = True
FRACONTRA.Visible = True
txtContra.SetFocus
End Sub
Private Sub procesaescala()
'SE TOMA LA CLAVE DEL PROVEEDOR ACTIVO Y SE APLICA UN 2 % MAS A LA ESCALA
'SE PRETENDE HACER CON SENTENCIA SQL
'SE PIDE CONTRASEÑA
v = MsgBox("Esta seguro de iniciar el proceso, Recuerde que solo se debe ejecutar una vez ", vbYesNo)
If v = vbNo Then
   Exit Sub
End If
Dim CLAVEPROD As String
Dim PAQUETES As Integer
Dim PRECIO1B As Double
Dim PRECIO2B As Double
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim rsesc As ADODB.Recordset
Set rsesc = New ADODB.Recordset
Dim rsprecio As ADODB.Recordset
Set rsprecio = New ADODB.Recordset
CADENA = "SELECT * FROM TFPRODUC  WHERE  CLAPROVE = '" & Trim(Me.TxtProv.Text) & "'  order by descripc "
rs.Source = CADENA
rs.LockType = adLockOptimistic
'rs.LockType = adLockBatchOptimistic
rs.CursorType = adOpenDynamic
rs.ActiveConnection = cCadConex
rs.Open
If rs.EOF Then
   MsgBox "No existen Productos de Este Proveedor..."
   rs.Close
   Exit Sub
End If
While Not rs.EOF
              
             If rs!MEDMAY = 0 Then
                CLAVEPROD = Trim(rs!CONSEC)
              '  If CLAVEPROD = "1004817" Then
              '     MsgBox "HOLA"
              '  End If
                costo = rs!PRECOSTO
                PAQUETES = rs!PAQUETES
                CADENA = "SELECT * FROM MARGEN  WHERE  PRODUCTO = '" & Trim(CLAVEPROD) & "'"
                rsesc.Source = CADENA
                rsesc.LockType = adLockOptimistic
                rsesc.CursorType = adOpenDynamic
                rsesc.ActiveConnection = cCadConex
                rsesc.Open
                'SE LE AUMENTA UN 2 PORCIENTO MAS
                If Not rsesc.EOF Then
                            escala1 = rsesc!escala1 + 2
                            escala2 = rsesc!escala2 + 2
                            rsesc!escala1 = escala1
                            rsesc!escala2 = escala2
                            rsesc.Update
                End If
                rsesc.Close
                'CALCULO DEL NUEVO PRECIO PARA ESCALAS 1 Y 2
                Dim npreciopaso As Double
                Dim NPRECIOPASO2 As Double
                npreciopaso = ((costo + (costo * escala1) / 100)) / PAQUETES
                PRECIO1B = calprecio(npreciopaso)
                NPRECIOPASO2 = ((costo + (costo * escala2) / 100))
                PRECIO2B = calprecio(NPRECIOPASO2)
                'SE GRABAN LOS PRECIOS
                CADENA = "SELECT * FROM PREPROD  WHERE  PRECLAVE = '" & Trim(CLAVEPROD) & "'"
                rsesc.Source = CADENA
                rsesc.ActiveConnection = cCadConex
                rsesc.LockType = adLockOptimistic
                rsesc.CursorType = adOpenDynamic
                rsesc.LockType = adLockBatchOptimistic
                rsesc.Open
                If Not rsesc.EOF Then
                    rsesc!precio1ant = rsesc!precio1
                    rsesc!precio2ant = rsesc!PRECIO2
                    rsesc!precio1 = PRECIO1B
                    rsesc!PRECIO2 = PRECIO2B
                    rsesc.Update
                End If
                rsesc.Close
           rs!fecact = date
           'rs!usuario = "ESPECIAL"
           rs.Update
           End If
           'Call CHECAMMAYOREO(CLAVEPROD, PAQUETES, PRECIO2B)
           rs.MoveNext
'           Adoprod.Recordset.MoveNext
Wend
rs.Close
MsgBox "PROCESO TERMINADO...", vbInformation
End Sub

Private Sub CHECAMMAYOREO(CLAVEPROD1 As String, PAQUETES As Integer, PRECIO As Double)
paq = PAQUETES
'SE DEBE HACER LA MULTIPLICACION POR EL PRECIO DE AUTOSERVICIO MAYOREO
'costomm = Round(PRECIO * 0.9, 2) / paq
costomm = Round(PRECIO * 1.02, 2) / paq
'MsgBox costomm
CADENA = "update tfproduc set costocaj =  round(" & costomm & " *  paquetes,1), precosto = round( " & costomm & " * paquetes,1), fecact = '" & date & "' where medmay = '" & Trim(CLAVEPROD1) & "'"
'MsgBox cadena
cn.Execute CADENA
'y tambien la parte precios preprod
pretemp = costomm
CADENA = "update preprod set precio1 = precosto , precio2 = precosto, precio3 = precosto, precio4 = precosto from tfproduc,preprod where medmay = '" & Trim(CLAVEPROD1) & "' and consec = preclave "
'MsgBox cadena
cn.Execute CADENA
End Sub

Private Function calprecio(npreciopaso As Double) As Double
PRECIO = 0
preciobase = npreciopaso
COMPARA1 = Int(npreciopaso)
COMPARA2 = npreciopaso - COMPARA1
If preciobase <= 20 Then
    If COMPARA2 > 0 Then
            If COMPARA2 <= 0.1 Then
                  PRECIO = COMPARA1 + 0.1
            ElseIf COMPARA2 <= 0.2 Then
                  PRECIO = COMPARA1 + 0.2
            ElseIf COMPARA2 <= 0.3 Then
                  PRECIO = COMPARA1 + 0.3
            ElseIf COMPARA2 <= 0.4 Then
                  PRECIO = COMPARA1 + 0.4
            ElseIf COMPARA2 <= 0.5 Then
                  PRECIO = COMPARA1 + 0.5
            ElseIf COMPARA2 <= 0.6 Then
                  PRECIO = COMPARA1 + 0.6
            ElseIf COMPARA2 <= 0.7 Then
                   PRECIO = COMPARA1 + 0.7
            ElseIf COMPARA2 <= 0.8 Then
                   PRECIO = COMPARA1 + 0.8
            ElseIf COMPARA2 <= 0.9 Then
                   PRECIO = COMPARA1 + 0.9
            Else
                  PRECIO = COMPARA1 + 1
            End If
   Else
           PRECIO = COMPARA1 'si el precio es un entero
   End If
Else
    If COMPARA2 <= 0.5 Then
           PRECIO = COMPARA1 + 0.5
    Else
           PRECIO = COMPARA1 + 1
    End If
End If
calprecio = PRECIO
End Function

Private Sub cmdprecios_Click()
 fmenu.cr1.WindowTitle = Strtitle
    strform1 = "{TFPRODUC.CLAPROVE} = '" & Trim(TxtProv.Text) & "'"
    fmenu.cr1.ReportFileName = App.Path & "\prodpre.rpt"
    fmenu.cr1.Formulas(0) = "formula1 = (" & strform1 & ")"
    fmenu.cr1.Connect = strconnect
    fmenu.cr1.WindowState = crptMaximized
    fmenu.cr1.Action = 1
End Sub

Private Sub CmdPREVER_Click()
'On Error GoTo Error:
Dim strform1 As String
cr1.Connect = strconnect
cr1.WindowTitle = "Hoja de catálogo de " & cmbProv.Text
If Trim(TxtProv.Text) <> "" Then
    strform1 = "{TFPRODUC.CLAPROVE} = '" & Trim(TxtProv.Text) & "' AND {TFPRODUC.activo} = 1"
    If Trim(Txtfamilia.Text) <> "" Then
        strform1 = strform1 + " AND {TFPRODUC.CLAFAMIL} = '" & Trim(Txtfamilia.Text) & "'"
        fmenu.cr1.Formulas(1) = "titulo = 'FAMILIA : " & Cmbfamilia.Text & "'"
    End If
    cr1.Formulas(0) = "formula1 = " & strform1
    'If TxtProv.Text = "C33" Then
    '   fmenu.CR1.ReportFileName = App.Path & "\hojacat1.rpt"
    'Else
     cr1.ReportFileName = App.Path & "\hojacat.rpt"
    'End If
    'MsgBox fmenu.CR1.Formulas(0)
    cr1.Action = 1
End If

Exit Sub
Error:
MsgBox " ¡¡ Los datos de los productos seleccionados no estan completos !!"
End Sub

Private Sub Command1_Click()
On Error GoTo Error:
Unload Me
Exit Sub
Error:
MsgBox Err.Description
End Sub
Private Sub Command2_Click()
' ir al form de precios
Dim reg
On Error GoTo Error:
If Adoprod.Recordset.RecordCount > 0 Then
reg = DGrprod.Bookmark
strcveprov = TxtProv.Text
strcveprod = DGrprod.Columns(0).Text
lpprov = True
lpprod = False
If tipotienda = 1 Or tipotienda = 4 Then
   frmprecios.Show
Else
   fnewprec.Show
End If
Adoprod.Refresh
DGrprod.Bookmark = reg
End If
Exit Sub
Error:
MsgBox "¡¡ Los productos de este proveedor no han sido clasificados !!"
End Sub

Private Sub Command3_Click()
' ir al form de proveedores
On Error GoTo Error:
If Trim(TxtProv.Text) <> "" Then
strcveprov = TxtProv.Text
lpprov = True
fprov.Show 1
AdoProv.Refresh
Cmbprov_Click
Else
   MsgBox "¡¡ No existe ningun proveedor seleccionado !!"
End If
Exit Sub
Error:
MsgBox "¡¡ Los productos de este proveedor no han sido clasificados !!"
Exit Sub
End Sub

Private Sub Command4_Click()
' presenta form de productos
Dim reg
On Error GoTo Error:
If Adoprod.Recordset.RecordCount > 0 Then
    reg = DGrprod.Bookmark
    lpprod = True
    strcveprov = TxtProv.Text
    strcveprod = DGrprod.Columns(0).Text
    frmprod.Show 1
    Adoprod.Refresh
    DGrprod.Bookmark = reg
End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub Command5_Click()
On Error GoTo Error:
If Adoprod.Recordset.RecordCount > 0 Then
    lpprov = True
    reg = DGrprod.Bookmark
    strcveprov = TxtProv.Text
    strcveprod = DGrprod.Columns(0).Text
    frmConvenio.Show 1
    Adoprod.Refresh
    DGrprod.Bookmark = reg
End If
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub Command6_Click()
On Error GoTo Error:

  cr1.Connect = cCadConex
  cr1.ReportFileName = App.Path & "\Exipro.rpt"
  If todos = True Then
     'EN EL CASO DE QUE SE TRATE DE TODOS LOS PRODUCTOS
      cr1.WindowTitle = "EXISTENCIAS TOTALES "
      cr1.Formulas(0) = "FORMSELEC = {TFPRODUC.CLAPROVE} <> '000' "
  Else
      cr1.WindowTitle = "EXISTENCIAS DE " & TxtProv.Text
      cr1.Formulas(0) = "FORMSELEC = {TFPRODUC.CLAPROVE} = '" & TxtProv.Text & "'"
  End If
  cr1.WindowState = crptMaximized
  cr1.Action = 1
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub DGrprod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim reg
On Error GoTo Error:
If Adoprod.Recordset.RecordCount > 0 Then
reg = DGrprod.Bookmark
strcveprov = TxtProv.Text
strcveprod = DGrprod.Columns(0).Text
lpprov = True
lpprod = False
If tipotienda = 1 Or tipotienda = 4 Then
   frmprecios.Show 1
Else
   fnewprec.Show
End If
Adoprod.Refresh
DGrprod.Bookmark = reg
End If
Exit Sub
Error:
MsgBox "¡¡ Los productos de este proveedor no han sido clasificados !!"
End If

End Sub

Private Sub Form_Activate()
  Unload frmAreaRecibo
End Sub

Private Sub Form_Load()
'On Error GoTo Error:

Adousuario.CommandType = adCmdText
Adousuario.ConnectionString = strconnect
Adousuario.CursorType = adOpenKeyset
Adousuario.LockType = adLockOptimistic
Adousuario.RecordSource = "SELECT * FROM usuarios"
Adousuario.Refresh
    
Call llenaprov
Exit Sub
Error:
MsgBox Err.Description

End Sub

Private Sub llenaprov()
'On Error GoTo Error:
   
    AdoProv.LockType = adLockOptimistic
    AdoProv.CursorType = adOpenKeyset
    AdoProv.CommandType = adCmdText
    activo = IIf(SoloAct, " and activo = 1", "")
    'Adoprov.RecordSource = "select * from catprov"
    If Nivel = "I" Then
       AdoProv.RecordSource = "SELECT * FROM usuarios U, catprov P WHERE p.comprador = u.login AND level1 = 'I'" & activo
    Else
       AdoProv.RecordSource = "SELECT * FROM usuarios U, catprov P WHERE p.comprador *= u.login AND level1 <> 'I'" & activo
    End If
    AdoProv.ConnectionString = strconnect
    AdoProv.Refresh
    
    cmbProv.Clear
    While AdoProv.Recordset.EOF = False
        If Not IsNull(AdoProv.Recordset!NOMPROVE) Then
            cmbProv.AddItem AdoProv.Recordset!NOMPROVE & "  [" & AdoProv.Recordset!prove & "]"
        End If
        AdoProv.Recordset.MoveNext
    Wend
    
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub llenaprod(cveprov As String)
On Error GoTo Error:
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim i As Integer
   cveprov = Trim(cveprov)
   
   Adoprod.CursorLocation = adUseServer
   Adoprod.LockType = adLockOptimistic
   Adoprod.CursorType = adOpenKeyset
   activo = IIf(SoloAct, "and tfproduc.activo = 1", "")
   Adoprod.RecordSource = "SELECT incant, consec,costototal,claprove,paquetes,barraspza,clafamil,linea,descripc,contenid,medida,decto1,decto2,decto3,decto4,financiero,efectivo,descprod.cajas,descprod.encajas,cargo1,cargo2,cargo3,cargo4,cargo5,flete,maniobras,TFPRODUC.activo," & _
                          "tfproduc.clavedelprov, descprod.plazopago,costocaj,precosto,precio1,precio1 * paquetes as precio_v,((( precio1 * paquetes )- costototal)/( precio1 * paquetes ))*100 as margen " & _
                          "FROM inventario, tfproduc,descprod,catprov,preprod WHERE consec=preclave " & activo & " AND consec*=producto and claprove = prove and  inprod = consec and claprove = '" & Trim(TxtProv.Text) & "' AND precio1 * paquetes > 0 ORDER BY descripc,contenid"
   Adoprod.ConnectionString = strconnect
   Adoprod.Refresh
   If Adoprod.Recordset.RecordCount > 0 Then
      Command2.Enabled = True  'precios
      Command4.Enabled = True   'productos
      Command5.Enabled = True   'condiciones de compra
      Command6.Enabled = True   'condiciones de compra
   Else
      Command2.Enabled = False
      Command4.Enabled = False
      Command5.Enabled = False
      Command6.Enabled = False
   End If
   'rs.Source = "select * from linprove,familias where clfamilia = sfclave and clprove = '" & Trim(cveprov) & "'"
   rs.Source = "select * from linprove,familias,departamento,catieps where clfamilia = fclave and clprove = '" & Trim(cveprov) & "' and fdepto = depclave and depclave = idepto and izona = '003'"
   rs.ActiveConnection = strconnect
   rs.Open

   Cmbfamilia.Clear
   While rs.EOF = False
       'If Not IsNull(rs.Fields!sfdescrip) Then
       If Not IsNull(rs.Fields!fdescrip) Then
           Cmbfamilia.AddItem rs.Fields!fdescrip & "  [" & rs.Fields!clfamilia & "]    Depto. (" & rs.Fields!depdescrip & ")"
       End If
       rs.MoveNext
   Wend
   rs.Close
Exit Sub
Error:
MsgBox Err.Description
End Sub





Private Sub Txtcar_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If
End Sub

Private Sub Txtdes_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
SendKeys "{BACKSPACE}"
Exit Sub
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmAreaRecibo.Show
End Sub

Private Sub Txtfamilia_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
Dim sdepto As String
Dim i As Integer
Dim familia As String
Dim lfamilia As Boolean
lfamilia = False
If KeyAscii = 13 Then
 Txtfamilia.Text = UCase(Txtfamilia.Text)
For i = 0 To Cmbfamilia.ListCount - 1
     N = InStr(1, Cmbfamilia.List(i), "[")
     familia = Mid(Cmbfamilia.List(i), N + 1, 3)
     If familia = Trim(Txtfamilia.Text) Then
       lfamilia = True
       Cmbfamilia.ListIndex = i
       N = InStr(1, Cmbfamilia.List(i), "(")
       sdepto = Mid(Cmbfamilia.List(i), N + 1, Len(Cmbfamilia.List(i)) - N - 1)
       Txtdepto.Text = sdepto
       Exit For
     End If
Next
If lfamilia Then
    Call Cmbfamilia_Click
Else
    MsgBox "La Clave no existe ... "
    Cmbfamilia.SetFocus
End If

End If
Exit Sub
Error:
MsgBox Err.Description
End Sub



Private Sub Txtfamilia_LostFocus()
Txtfamilia_KeyPress 13
End Sub


Private Sub TxtProv_KeyPress(KeyAscii As Integer)
On Error GoTo Error:
Dim i As Integer
Dim prov As String
Dim lprov As Boolean
lprov = False
If KeyAscii = 13 Then
TxtProv.Text = UCase(TxtProv.Text)
For i = 0 To cmbProv.ListCount - 1
     N = InStr(1, cmbProv.List(i), "[")
     prov = Mid(cmbProv.List(i), N + 1, Len(cmbProv.List(i)) - N - 1)
     If Trim(prov) = Trim(TxtProv.Text) Then
       lprov = True
       cmbProv.ListIndex = i
       Txtdepto.Text = ""
       Exit For
     End If
Next
If lprov Then
    Call Cmbprov_Click
Else
    'MsgBox "La Clave no existe ... "
    cmbProv.SetFocus
End If

End If
Exit Sub
Error:
MsgBox Err.Description
End Sub

Private Sub TxtProv_LostFocus()
TxtProv_KeyPress 13
End Sub
