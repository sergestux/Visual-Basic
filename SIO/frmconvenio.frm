VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConvenio 
   Caption         =   "Condiciones de compra"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   ForeColor       =   &H00000000&
   Icon            =   "frmconvenio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adoprod 
      Height          =   375
      Left            =   120
      Top             =   3840
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
   Begin VB.Frame Frame1 
      Caption         =   "Promociones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   23
      Top             =   5880
      Width           =   3015
      Begin MSMask.MaskEdBox Mskencajas 
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskcajas 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "En la compra de"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Num. Cajas"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Descuentos :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   21
      Top             =   4560
      Width           =   4815
      Begin VB.TextBox Txtdes 
         Height          =   375
         Index           =   6
         Left            =   2040
         MaxLength       =   7
         TabIndex        =   48
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Txtdes 
         Height          =   375
         Index           =   5
         Left            =   3480
         MaxLength       =   7
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Txtdes 
         Height          =   375
         Index           =   4
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Txtdes 
         Height          =   375
         Index           =   3
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Txtdes 
         Height          =   375
         Index           =   2
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Txtdes 
         Height          =   375
         Index           =   1
         Left            =   600
         MaxLength       =   5
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Txtdes 
         Height          =   375
         Index           =   0
         Left            =   120
         MaxLength       =   5
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "    1%       2%      3%       4%        5%       Efectivo  $   Financiero %"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cargos :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   19
      Top             =   4440
      Width           =   5055
      Begin VB.TextBox Txtcar 
         Height          =   375
         Index           =   6
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Txtcar 
         Height          =   375
         Index           =   5
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Txtcar 
         Height          =   375
         Index           =   1
         Left            =   600
         MaxLength       =   5
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Txtcar 
         Height          =   375
         Index           =   0
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   480
         Width           =   510
      End
      Begin VB.TextBox Txtcar 
         Height          =   375
         Index           =   4
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Txtcar 
         Height          =   375
         Index           =   3
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Txtcar 
         Height          =   375
         Index           =   2
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "       1%      2%   IEPS%  IVA%   Efectivo $     $ Flete         $Maniobras      "
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6855
      Left            =   360
      TabIndex        =   24
      Top             =   360
      Width           =   12120
      Begin VB.Frame Frame5 
         Caption         =   "Plazo de Pago "
         Height          =   975
         Left            =   3600
         TabIndex        =   45
         Top             =   5640
         Width           =   1935
         Begin VB.TextBox Txtplazopago 
            Height          =   375
            Left            =   480
            MaxLength       =   3
            TabIndex        =   16
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Dias"
            Height          =   375
            Left            =   1440
            TabIndex        =   46
            Top             =   600
            Width           =   495
         End
      End
      Begin MSMask.MaskEdBox Mskcompra 
         Height          =   375
         Left            =   7800
         TabIndex        =   42
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskcosto 
         Height          =   375
         Left            =   2400
         TabIndex        =   0
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cerrar"
         Height          =   495
         Index           =   1
         Left            =   8520
         Picture         =   "frmconvenio.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Regresar a la pantalla de pedidos"
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CommandButton Cmdaccion 
         Caption         =   "&Grabar"
         Height          =   495
         Index           =   0
         Left            =   6480
         Picture         =   "frmconvenio.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Actualizar tabla de precios"
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Precio Compra :"
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
         Left            =   6360
         TabIndex        =   41
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Costo  Por Caja :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   40
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Lblprod 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Index           =   7
         Left            =   7200
         TabIndex        =   39
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Lblprod 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Index           =   6
         Left            =   5400
         TabIndex        =   38
         Top             =   2760
         Width           =   4575
      End
      Begin VB.Label Lblprod 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Index           =   5
         Left            =   5400
         TabIndex        =   37
         Top             =   2280
         Width           =   4575
      End
      Begin VB.Label Lbletiq 
         BackStyle       =   0  'Transparent
         Caption         =   "Familia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   6
         Left            =   4680
         TabIndex        =   36
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Lbletiq 
         BackStyle       =   0  'Transparent
         Caption         =   "Linea :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   35
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Lblprod 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Index           =   4
         Left            =   6120
         TabIndex        =   34
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Lbletiq 
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   4
         Left            =   4680
         TabIndex        =   33
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Lblprod 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Index           =   3
         Left            =   2400
         TabIndex        =   32
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Lblprod 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   31
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Lbletiq 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Lbletiq 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo de Barras : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   29
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Lbletiq 
         BackStyle       =   0  'Transparent
         Caption         =   "Presentación : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Lbletiq 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Lblprod 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   26
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label Lblprod 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808000&
         FillColor       =   &H00FFFFC0&
         Height          =   2295
         Left            =   360
         Top             =   1200
         Width           =   9615
      End
   End
   Begin VB.Label Label3 
      Caption         =   "CONDICIONES PARA REALIZAR LOS PEDIDOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   47
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lnuevo As Boolean
Private lClas As Boolean

Private Sub Cmdaccion_Click(Index As Integer)
'On Error GoTo eRROR:
Dim rsttemp As ADODB.Recordset
 'Procedimiento que actualiza el precio de compra en base a los cargos y desctos. especificados
 llenagrid
 'Convierto todos los precios a historicos
 cn.Execute "UPDATE descprod SET situacion = 0  WHERE producto = " & Trim(Adoprod.Recordset!CONSEC) & ""
 Set rsttemp = New ADODB.Recordset
 rsttemp.Open "SELECT * FROM Descprod WHERE Producto = '" & Trim(Adoprod.Recordset!CONSEC) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
 'CHECAR EL FINANCIERO ES DES(5), EL EFECTIVO ES DES(4) Y EL DES5 ES DES(6)
 If rsttemp.BOF And rsttemp.EOF Then
  cn.Execute "INSERT INTO descprod (proveedor,producto,decto1,decto2,decto3,decto4,financiero,efectivo,cargo1,cargo2,cargo3,cargo4,cargo5,cajas,encajas,costo,preciolista,flete,maniobras,plazopago,situacion, fechaact,decto5)" & _
               " VALUES ('" & Adoprod.Recordset!claprove & "','" & Adoprod.Recordset!CONSEC & "'," & Val(Txtdes(0).Text) & "," & Val(Txtdes(1).Text) & "," & _
                Val(Txtdes(2).Text) & ", " & Val(Txtdes(3).Text) & "," & Val(Txtdes(5).Text) & "," & Val(Txtdes(4).Text) & _
               "," & Val(Txtcar(0).Text) & "," & Val(Txtcar(1).Text) & "," & _
                Val(Txtcar(3).Text) & "," & Val(Txtcar(2).Text) & "," & Val(Txtcar(4).Text) & "," & _
                Val(Mskcajas.Text) & "," & Val(Mskencajas.Text) & "," & Val(Mskcompra.Text) & "," & Val(Mskcosto.Text) & "," & Val(Txtcar(5).Text) & "," & Val(Txtcar(6).Text) & "," & Val(Txtplazopago.Text) & _
                ",1" & ",'" & date + Time & "'," & Val(Txtdes(6).Text) & ")"
Else
    CADENA = "UPDATE descprod SET proveedor = '" & Adoprod.Recordset!claprove & "',producto = '" & Adoprod.Recordset!CONSEC & "',decto1 = " & Val(Txtdes(0).Text) & ",decto2 =" & Val(Txtdes(1).Text) & ",decto3 =" & Val(Txtdes(2).Text) & "," & _
              "decto4 =" & Val(Txtdes(3).Text) & ",financiero=" & Val(Txtdes(5).Text) & ",efectivo=" & Val(Txtdes(4).Text) & ",cargo1 =" & Val(Txtcar(0).Text) & ",cargo2=" & Val(Txtcar(1).Text) & ",cargo3=" & Val(Txtcar(3).Text) & ",cargo4=" & Val(Txtcar(2).Text) & "," & _
              "cargo5=" & Val(Txtcar(4).Text) & ",cajas=" & Val(Mskcajas.Text) & ",encajas=" & Val(Mskencajas.Text) & ",costo=" & Val(Mskcompra.Text) & ",preciolista=" & Val(Mskcosto.Text) & ",situacion = 1,fechaact = '" & date + Time & "', flete =" & Val(Txtcar(5).Text) & ",maniobras =" & Val(Txtcar(6).Text) & ",plazopago =" & Val(Txtplazopago.Text) & _
              ", decto5=" & Val(Txtdes(6).Text) & _
              " WHERE Producto = '" & Adoprod.Recordset!CONSEC & "'"
     'MsgBox cadena
    cn.Execute CADENA
End If
'AQUI NO SE DEBE ACTUALIZAR EL TFPRODUC
'cn.Execute "UPDATE TFPRODUC SET ACTUALIZADO = 1 WHERE CONSEC = '" & Adoprod.Recordset!consec & "'"
Unload Me
Exit Sub
eRROR:
MsgBox Err.Description
End Sub

Private Sub Command1_Click(Index As Integer)
  Unload Me
End Sub

Private Sub Form_Activate()
 If Not lClas Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo eRROR:
   Adoprod.CursorLocation = adUseServer
   Adoprod.LockType = adLockOptimistic
   Adoprod.CursorType = adOpenKeyset
   Adoprod.RecordSource = "SELECT CLAPROVE,cajas,encajas,consec,precosto,TFPRODUC.ACTIVO,barraspza,contenid,medida,descripc,costocaj,paquetes,nomprove,sfdescrip,fdescrip,depdescrip,iporcentaje " & _
        "FROM tfproduc,catprov,familias,lineas,departamento,catieps " & _
        "WHERE claprove = prove AND linea = sfclave AND sffamilia = fclave AND fdepto = depclave AND TFPRODUC.Consec = '" & Trim(strcveprod) & "'"
   Adoprod.ConnectionString = strconnect
   Adoprod.Refresh
   lClas = True
   If Adoprod.Recordset.BOF And Adoprod.Recordset.EOF Then
      MsgBox "EL ARTICULO ESPECIFICADO NO ESTA CLASIFICADO CORRECTAMENTE", vbCritical
      lClas = False
      Exit Sub
   End If
   Call asigna
Exit Sub

eRROR:
MsgBox Err.Description
End Sub

Private Sub Mskcajas_GotFocus()
  Mskcajas.SelLength = Len(Mskcajas.Text)
End Sub

Private Sub Mskcajas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Mskcosto_GotFocus()
  Mskcosto.SelLength = Len(Mskcosto.Text)
End Sub

Private Sub Mskcosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Mskcosto_LostFocus()
llenagrid
End Sub


Private Sub Mskencajas_GotFocus()
  Mskencajas.SelLength = Len(Mskencajas.Text)
End Sub

Private Sub Mskencajas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
End Sub

Private Sub Mskencajas_LostFocus()
  llenagrid
End Sub

Private Sub Txtcar_GotFocus(Index As Integer)
Txtcar(Index).SelLength = Len(Txtcar(Index).Text)
End Sub

Private Sub Txtcar_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
   SendKeys "{BACKSPACE}"
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
Exit Sub
End Sub


Private Sub Txtcar_LostFocus(Index As Integer)
  llenagrid
End Sub

Private Sub Txtdes_GotFocus(Index As Integer)
Txtdes(Index).SelLength = Len(Txtdes(Index).Text)
End Sub

Private Sub Txtdes_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo eRROR:

If (KeyAscii < 46 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
   SendKeys "{BACKSPACE}"
   Exit Sub
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   SendKeys vbTab
End If
Exit Sub

eRROR:
MsgBox Err.Description
End Sub
Private Sub llenagrid()
On Error GoTo eRROR:
Dim npreciodes As Double
Dim nprecio As Double
Dim i As Integer
    nprecio = Val(Mskcosto.Text)
      'calcula cargos %
    For i = 0 To 3
        nprecio = nprecio + (nprecio * (Val(Txtcar(i).Text) / 100))
      '  MsgBox Txtcar(I).Text
    Next i
    For i = 4 To 6
        nprecio = nprecio + Val(Txtcar(i).Text)
      '  MsgBox Txtcar(I).Text
    Next i

    
      'cargo efectivo
    'nprecio = nprecio + Val(Txtcar(4).Text)
     
     'calcula descuentos %
    nprecio = Round(nprecio, 2)
    'OJO AQUI SE DEBE METER LA CONDICION DE LOS CENTAVOS
    
    For i = 0 To 3
         nprecio = nprecio - (nprecio * (Val(Txtdes(i).Text) / 100))
    Next i
    'EL DESCUENTO NUMERO 5
    nprecio = nprecio - (nprecio * (Val(Txtdes(6).Text) / 100))
    'descuento efectivo
    'nprecio = nprecio - (nprecio * (Val(Txtdes(4).Text) / 100))
    
    nprecio = nprecio - Val(Txtdes(4).Text)
        
    'descuento financiero
    nprecio = nprecio - (nprecio * (Val(Txtdes(5).Text) / 100))
    
    Mskcompra.Text = Round(nprecio, 2)
       
Exit Sub
eRROR:
MsgBox Err.Description



End Sub

Sub asigna()
Dim rs As ADODB.Recordset
On Error GoTo eRROR:
    Lblprod(1).Caption = Adoprod.Recordset!DESCRIPC
    Lblprod(3).Caption = Adoprod.Recordset!barraspza
    Lblprod(6).Caption = Adoprod.Recordset!sfdescrip
    Lblprod(4).Caption = Adoprod.Recordset!depdescrip
    Lblprod(5).Caption = Adoprod.Recordset!fdescrip
    Lblprod(7).Caption = IIf(Adoprod.Recordset!activo = True, "ACTIVO", "CANCELADO")
    Lblprod(7).ForeColor = IIf(Adoprod.Recordset!activo = True, &HFFFFFF, &HFF&)
    
    
    Lblprod(2).Caption = Trim(Str(Adoprod.Recordset!PAQUETES)) + " X 1          " + Trim(Str(Adoprod.Recordset!Contenid)) + " " + Adoprod.Recordset!medida
    Lblprod(0).Caption = Adoprod.Recordset!CONSEC
    Frame2.Visible = True
   ' moy
    Mskcajas.Text = Adoprod.Recordset!cajas
    Mskencajas.Text = Adoprod.Recordset!encajas
    Mskcosto.Text = Adoprod.Recordset!costocaj
    Me.Mskcompra.Text = 0
        
    For i = 0 To 5
    Txtdes(i).Text = 0
    Next
    
    For i = 0 To 6
    Txtcar(i).Text = 0
    Next
   Set rs = New ADODB.Recordset
   rs.LockType = adLockOptimistic
   rs.CursorType = adOpenKeyset
   rs.Source = "SELECT * FROM DESCPROD WHERE  PRODUCTO = '" & Trim(Adoprod.Recordset!CONSEC) & "'"

   rs.ActiveConnection = cn
   rs.Open
   If rs.RecordCount > 0 Then
    Txtdes(0).Text = rs.Fields!decto1
    Txtdes(1).Text = rs.Fields!decto2
    Txtdes(2).Text = rs.Fields!decto3
    Txtdes(3).Text = rs.Fields!decto4
    Txtdes(6).Text = rs.Fields!decto5
    'ERROR
    Txtdes(5).Text = rs.Fields!financiero
    Txtdes(4).Text = rs.Fields!efectivo
    'FIN DEL ERRROR CHECAR
    Mskcajas.Text = rs.Fields!cajas
    Mskencajas.Text = rs.Fields!encajas
    Mskcosto.Text = rs.Fields!preciolista
    Me.Mskcompra.Text = rs.Fields!costo
    Txtcar(0).Text = rs.Fields!cargo1
    Txtcar(1).Text = rs.Fields!cargo2
    Txtcar(2).Text = rs.Fields!cargo4
    Txtcar(3).Text = rs.Fields!cargo3
    Txtcar(4).Text = rs.Fields!cargo5
    Txtcar(5).Text = rs.Fields!flete
    Txtcar(6).Text = rs.Fields!maniobras
    Txtplazopago.Text = rs.Fields!plazopago
    lnuevo = False
   End If
   Cmdaccion(0).Enabled = (tipotienda = 1 Or tipotienda = 4)
Exit Sub
eRROR:
MsgBox Err.Description

End Sub


Private Sub Txtdes_LostFocus(Index As Integer)
  llenagrid
End Sub
