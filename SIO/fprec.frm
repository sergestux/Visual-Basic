VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form fnewprec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios de Productos..."
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10050
   ClipControls    =   0   'False
   Icon            =   "fprec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   10050
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   4860
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "                                                                                             Para salir presione la tecla Esc"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9735
      Begin VB.Label lblcodbarra 
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   4935
      End
      Begin VB.Label lblcod 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblpres 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label lblprod 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   9375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   9735
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   2280
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   2280
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   7200
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   7200
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   7200
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2280
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   -120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "1/2 Caja Bodega"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "1/2 Caja Envío"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Caja  Bodega"
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Caja Intermedio"
         Height          =   255
         Left            =   5520
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Caja crédito y/o envío"
         Height          =   495
         Left            =   5520
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Pieza Autoservicio"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Costo"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblconsec 
         Caption         =   "."
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "fnewprec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private descr As String
Private codigo As String
Private precio1
Private medida As String

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Set adorsTemp = New Recordset
adorsTemp.LockType = adLockOptimistic
adorsTemp.CursorType = adOpenKeyset
adorsTemp.ActiveConnection = cCadConex 'cn
'MsgBox strcveprod
If Sql Then
  adorsTemp.Open "select contenid,medida,FECACT,barraspza,descripc,str(paquetes) + ' x ' + ltrim(str(contenid,10,3)) + ' ' + medida as medida1,consec,precosto,precio1,precio2,precio3,precio4,PRECIO5, Precio6 from tfproduc,preprod where consec = preclave and consec = '" & Trim(strcveprod) & "'"
Else
  adorsTemp.Open "select contenid,medida,FECACT,barraspza,descripc,LTRIM(STR(T.paquetes)) + ' X ' + lTrim(str(T.contenid)) + space(2) + t.medida As MEDIDA1 ,consec,precosto,precio1,precio2,precio3,precio4,precio5,precio6 from tfproduc T,preprod where consec = preclave and consec = '" & Trim(strcveprod) & "'"
End If
If Not adorsTemp.EOF Then
    Text1.Text = Format(adorsTemp!PRECOSTO, "$#,###,###.00")
    Text2(0).Text = Format(adorsTemp!precio1, "$#,###,###.00")
    Text2(1).Text = Format(adorsTemp!PRECIO2, "$#,###,###.00")
    Text2(2).Text = Format(adorsTemp!PRECIO3, "$#,###,###.00")
    Text2(3).Text = Format(adorsTemp!precio4, "$#,###,###.00")
    Text2(4).Text = Format(adorsTemp!precio5, "$#,###,###.00")
    Text2(5).Text = Format(adorsTemp!precio6, "$#,###,###.00")
End If
lblprod.Caption = "Descripcion : " & adorsTemp!descripc
lblpres.Caption = "Presentación: " & adorsTemp!medida1
lblcodbarra.Caption = "Código de barras: " & adorsTemp!barraspza
lblcod.Caption = "Fecha de actualización:   " & adorsTemp!fecact
descr = adorsTemp!descripc
codigo = adorsTemp!barraspza
medida = adorsTemp!CONTENID & " " & adorsTemp!medida
precio1 = adorsTemp!precio1
adorsTemp.Close
'Text1.Visible = Not Sql
'Label1.Visible = Not Sql
End Sub

