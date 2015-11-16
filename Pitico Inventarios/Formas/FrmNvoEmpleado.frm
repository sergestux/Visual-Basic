VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmNvoEmpleado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo Empleado"
   ClientHeight    =   3240
   ClientLeft      =   2760
   ClientTop       =   4200
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1920
      MouseIcon       =   "FrmNvoEmpleado.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Da de alta al nuevo empleado"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4440
      MouseIcon       =   "FrmNvoEmpleado.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Regresa a la opcion anterior"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame FrameResponsable 
      Caption         =   "Empleados"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.TextBox TxtNombre 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         ToolTipText     =   "Nombre del empleado"
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox TxtClave 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         ToolTipText     =   "Clave del empleado"
         Top             =   480
         Width           =   735
      End
      Begin MSDataListLib.DataCombo CmbTienda 
         Bindings        =   "FrmNvoEmpleado.frx":0614
         DataField       =   "CLAVE"
         DataMember      =   "Tiendas"
         DataSource      =   "DE1"
         Height          =   315
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Tienda en que labora el empleado"
         Top             =   1320
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "CLAVE"
         Text            =   "DataCombo1"
         Object.DataMember      =   "Tiendas"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   -1440
         TabIndex        =   6
         Top             =   1080
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmNvoEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'Las declaraciones de variables seran obligatorias

Private Sub CmbTienda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdAceptar.SetFocus
End Sub
'Agrega un nuevo empleado
Private Sub CmdAceptar_Click()
Dim SQL As String
On Error GoTo Error
    
    SQL = "Insert Into EMPLEADOS Values ('" & TxtClave & "','" & TxtNombre & "','" & CmbTienda.BoundText & "')"
    Conn.BeginTrans
    Conn.Execute SQL
    IMPRIME "Empleado Registrado correctamente"
    Conn.CommitTrans
    
    If MsgBox("Desea agregar otro empleado", vbQuestion + vbYesNo) = vbYes Then
        TxtClave = ""
        TxtNombre = ""
        TxtClave.SetFocus
    Else
        FrmEmpleados.TxtClave = TxtClave
        FrmEmpleados.TxtNombre = TxtNombre
        FrmEmpleados.CmbTiendas.BoundText = CmbTienda.BoundText
        
        FrmEmpleados.CmbFuncion.SetFocus
        Unload Me
    End If
    Exit Sub
Error:
    Errores
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtClave <> "" Then TxtNombre.SetFocus
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
    If KeyAscii = 13 And TxtNombre <> "" Then CmbTienda.SetFocus
End Sub

Sub Errores()
    Select Case Err.Number
'        Case -2147467259        'Llave primaria duplicada
'            IMPRIME "La clave del empleado ya esta en uso"
'            TxtClave.SetFocus
        Case Else
            IMPRIME Err.Description
    End Select
    
    Conn.RollbackTrans  'Deshago la transaccion
    Exit Sub
End Sub
