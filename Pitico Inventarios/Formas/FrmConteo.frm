VERSION 5.00
Begin VB.Form FrmConteo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conteo"
   ClientHeight    =   4050
   ClientLeft      =   1095
   ClientTop       =   1935
   ClientWidth     =   9900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "BORRAR TABLAS"
      Height          =   495
      Left            =   7440
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CmdContar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame FrameConteo 
      Caption         =   "Conteo"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   9015
      Begin VB.TextBox TxtUbicacion 
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
         Left            =   3000
         TabIndex        =   4
         Text            =   "R01"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox TxtConteo 
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
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "1"
         ToolTipText     =   "Numero de conteo, P.ej. 2"
         Top             =   1200
         Width           =   300
      End
      Begin VB.TextBox TxtRespConteo 
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Text            =   "R Conteo"
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion"
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
         Left            =   1920
         TabIndex        =   7
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conteo"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable del Conteo"
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
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   2250
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "FrmConteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdContar_Click()
Dim SQL As String
On Error GoTo Error
    
    CONTEO = TxtConteo          'Actualizo la variable del conteo actual
    UBICACION = TxtUbicacion    'Actualizo la variable de la ubicacion actual
    INVENTARIO = TxtClaveTienda & Format(Date, "yymmdd")
    
    Select Case CONTEO
        Case 1  'Si es el primer conteo
        SQL = "Insert Into INVENTARIO Values ('" & INVENTARIO & "','" & Date & "','" & TIENDA & "','" & TxtRespInventario & "','"
        SQL = SQL & TxtRespConteo & "','','')"
            
        Case 2  'Si es el segundo conteo
            SQL = "Update Inventario Set Resp_Conteo2='" & TxtRespConteo & "' Where clave='" & INVENTARIO & "'"
            
        Case 3  'Si es el tercer conteo
            SQL = "Update Inventario Set Resp_Conteo3='" & TxtRespConteo & "' Where clave='" & INVENTARIO & "'"
    End Select
    
    'Debug.Print SQL
    Conn.BeginTrans
    Conn.Execute SQL
    Conn.CommitTrans
    'Barra.Panels(1).Text = "Tienda registrada correctamente"
    FrmInventaria.Show
    
    Exit Sub            'Salgo del procedimiento
    
Error:
    Conn.RollbackTrans      'SI hay un error se deshace la transacion
    Errores
    
End Sub

Sub Errores()
Dim MENSAJE As String
    Select Case Err.Number
        Case -2147467259        'Llave primaria duplicada
            MENSAJE = "Alguna clave ya esta en uso"
        Case Else
            MENSAJE = Err.Description
    End Select
    
    
    Beep
    Barra.Panels(1).Text = MENSAJE
    
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
