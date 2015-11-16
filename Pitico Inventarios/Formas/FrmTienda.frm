VERSION 5.00
Begin VB.Form FrmTienda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tienda"
   ClientHeight    =   3090
   ClientLeft      =   1230
   ClientTop       =   2175
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2760
      MouseIcon       =   "FrmTienda.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Da de alta la tienda actual"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   5400
      MouseIcon       =   "FrmTienda.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Regresa al menu principal"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame FrameTienda 
      Caption         =   "Tienda a Inventariar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      Begin VB.TextBox TxtMunicipio 
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Text            =   "TAPACHULA"
         ToolTipText     =   "Ubicacion de la tienda"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         ToolTipText     =   "Ubicacion de la tienda"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         ToolTipText     =   "Descripcion de la tienda, P. j. 'Pitico 1'"
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox TxtClaveTienda 
         Height          =   375
         Left            =   720
         MaxLength       =   5
         TabIndex        =   1
         ToolTipText     =   "Clave de tienda"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   2
         Left            =   5760
         TabIndex        =   11
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   450
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
Attribute VB_Name = "FrmTienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'Las declaraciones de variables seran obligatorias

Private Sub CmdAceptar_Click()
Dim SQL As String
On Error GoTo Error
    
    '**********************ACTUALIZO TIENDA*********************************
    SQL = "Insert Into TIENDAS Values ('" & TxtClaveTienda & "','" & TxtDescripcion & "','" & TxtDireccion & "','" & TxtMunicipio & "')"
    Conn.BeginTrans
    Conn.Execute SQL
    TIENDA = TxtClaveTienda     'Actualizo la Variable Tienda
    Conn.CommitTrans
    IMPRIME "Tienda registrada correctamente"
    
    If MsgBox("¿Desea agregar otra tienda?", vbQuestion + vbYesNo) = vbYes Then
        TxtClaveTienda = ""
        TxtDescripcion = ""
        TxtDireccion = ""
        TxtClaveTienda.SetFocus
    Else
        Unload Me
    End If
        
    Exit Sub            'Salgo del procedimiento
    
Error:
    Conn.RollbackTrans      'SI hay un error se deshace la transacion
    Errores
    
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Sub Errores()
    Select Case Err.Number
'        Case -2147467259        'Llave primaria duplicada
'            IMPRIME "La clave de la tienda ya esta en uso, use otra"
        Case Else
            IMPRIME Err.Description
    End Select
    
End Sub
Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
End Sub

Private Sub TxtMunicipio_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
End Sub
