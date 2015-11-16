VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcapinv 
   Caption         =   "Captura de Inventario..."
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstProd 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   6075
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8916
      _Version        =   393216
      Rows            =   50
      Cols            =   4
      WordWrap        =   -1  'True
   End
   Begin MSAdodcLib.Adodc adoprod 
      Height          =   495
      Left            =   4920
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      Caption         =   "Adodc2"
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
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   8535
      Begin VB.CommandButton cmdopcion 
         Caption         =   "&Salidas"
         Height          =   495
         Index           =   2
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "&Grabar"
         Height          =   495
         Index           =   1
         Left            =   4080
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "&Entradas"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc adoentrada 
      Height          =   495
      Left            =   0
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
   Begin VB.Frame Frmbonotes 
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.TextBox txtfecha 
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtfactura 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Factura"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmcapinv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_Click()

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
End Sub

Private Sub cmdopcion_Click(Index As Integer)
lstProd.Visible = True
lstProd.Enabled = True
End Sub

Private Sub Form_Load()
'Cargo los productos
adoprod.ConnectionString = cCadConex
adoprod.CommandType = adCmdText
adoprod.RecordSource = " select * from tfproduc"
adoprod.Refresh
lstProd.Clear
Do While Not adoprod.Recordset.EOF
nReg = nReg + 1
If Not (IsNull(adoprod.Recordset!MEDIDA)) And Not IsNull(adoprod.Recordset!PAQUETES) Then
    lstProd.AddItem adoprod.Recordset!Descripc + " ( " + adoprod.Recordset!MEDIDA + " ) " _
     + " [ " + adoprod.Recordset!consec + " ]"
End If
adoprod.Recordset.MoveNext
Loop

grid1.ColWidth(0) = 0 '1000 * 1
grid1.ColWidth(1) = 800
grid1.ColWidth(2) = 1000 * 8
grid1.ColWidth(3) = 1000 * 1.2
End Sub

Private Sub gridcaptura_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex = 0 Then
  'presentar el catalogo de productos
  lstProd.Enabled = True
  lstProd.Visible = True
End If
End Sub


Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  MsgBox "moy"
  lstProd.Visible = True
  lstProd.Enabled = True
End If
End Sub

Private Sub lstProd_KeyPress(KeyAscii As Integer)
'On Error Resume Next
Dim Cveprod As String
Dim n As Integer
    'Habilita el uso de teclado para asignar a la celda o Escape...
    Select Case KeyAscii
        Case vbKeyReturn
             n = InStr(1, lstProd.List(lstProd.ListIndex), "[")
             Cveprod = Mid(lstProd.List(lstProd.ListIndex), n + 1, Len(lstProd.List(lstProd.ListIndex)) - n - 1)
             lstProd.Visible = False
             grid1.AddItem (1)
            grid1.Col = 0
            grid1.Text = Cveprod
            grid1.Col = 1
            grid1.Text = lstProd.List(lstProd.ListIndex)
             'llenaproducto (Cveprod)
        Case vbKeyEscape
             'LblProdAgr.Caption = ""
             lstProd.Visible = False
             'DbgrdDetTraAbi.SetFocus   'Para que se posicione en la columna de cajas
    End Select
End Sub

