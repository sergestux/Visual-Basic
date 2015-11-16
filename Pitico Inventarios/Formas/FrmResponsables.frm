VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmResponsables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Responsables"
   ClientHeight    =   5055
   ClientLeft      =   1065
   ClientTop       =   2475
   ClientWidth     =   9810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "BORRAR TABLAS"
      Height          =   495
      Left            =   7440
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame FrameTienda 
      Caption         =   "Responsable del Inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin MSDataListLib.DataCombo CmbRespInventario 
         Bindings        =   "FrmResponsables.frx":0000
         DataField       =   "NOMBRE"
         DataMember      =   "RespInventarios"
         DataSource      =   "DE1"
         Height          =   315
         Left            =   3120
         TabIndex        =   14
         Top             =   2640
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NOMBRE"
         Text            =   "DataCombo2"
         Object.DataMember      =   "RespInventarios"
      End
      Begin MSDataListLib.DataCombo CmbRespTienda 
         Bindings        =   "FrmResponsables.frx":0022
         DataField       =   "CLAVE"
         DataMember      =   "RespTiendas"
         DataSource      =   "DE1"
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   1920
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NOMBRE"
         BoundColumn     =   "CLAVE"
         Text            =   "DataCombo1"
         Object.DataMember      =   "RespTiendas"
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Text            =   "Conocida"
         Top             =   1200
         Width           =   6615
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Text            =   "Pitico 01"
         ToolTipText     =   "Descripcion de la tienda, P. j. 'Pitico 1'"
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox TxtClaveTienda 
         Height          =   375
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "P01"
         ToolTipText     =   "Clave de tienda, "
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable de Tienda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   2040
         Width           =   2250
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable del Inventario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   8
         Top             =   2760
         Width           =   2520
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   -1440
         TabIndex        =   4
         Top             =   1080
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmResponsables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
Dim SQL As String
On Error GoTo Error

'CmbFuncion.BoundText
    
    CONTEO = TxtConteo          'Actualizo la variable del conteo actual
    UBICACION = TxtUbicacion    'Actualizo la variable de la ubicacion actual
    INVENTARIO = TxtClaveTienda & Format(Date, "yymmdd")
    
    '****************ACTUALIZO TIENDA**********************
    SQL = "Insert Into TIENDAS Values ('" & TxtClaveTienda & "','" & TxtDescripcion & "','" & TxtDireccion & "')"
    Conn.BeginTrans
    Conn.Execute SQL
    TIENDA = TxtClaveTienda     'Actualizo la Variable Tienda
    Conn.CommitTrans
    Beep
    Barra.Panels(1).Text = "Tienda registrada correctamente"
    
    FrmConteo.Show
    
    Exit Sub            'Salgo del procedimiento
    
Error:
    Conn.RollbackTrans      'SI hay un error se deshace la transacion
    Errores
    
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Sub Errores()
Dim MENSAJE As String
    Select Case Err.Number
        Case -2147467259        'Llave primaria duplicada
            MENSAJE = "La clave de la tienda ya esta en uso, use otra"
        Case Else
            MENSAJE = Err.Description
    End Select
    
    Beep
    FrmPrincipal.Barra.Panels(1).Text = MENSAJE
    
End Sub


Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    Conn.Execute "delete from conteo"
    Conn.Execute "delete from inventario"
    Conn.Execute "delete from tienda"
    Conn.Execute "delete from responsable"
End Sub

Private Sub Form_Activate()
    If CONTEO > 1 Then
        TxtClaveTienda.Enabled = False
        TxtDescripcion.Enabled = False
        TxtDireccion.Enabled = False
        TxtRespInventario.Enabled = False
        TxtRespTienda.Enabled = False
    End If
End Sub


'Private Sub TxtConteo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then TxtUbicacion.SetFocus
'End Sub

Private Sub TxtConteo_Validate(Cancel As Boolean)
    If Val(TxtConteo) < 1 Or Val(TxtConteo) > 3 Then
        Beep
        Barra.Panels(1).Text = "Escriba 1, 2 o 3, Segun el conteo correspondiente"
        Cancel = True     'Dejo que Pierda el enfoque
    Else
        Cancel = False  'Dejo que Pierda el enfoque
        TxtUbicacion.SetFocus
    End If
End Sub

Private Sub TxtUbicacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdContar.SetFocus
End Sub
