VERSION 5.00
Begin VB.Form FrmFunciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Funciones"
   ClientHeight    =   1815
   ClientLeft      =   -45
   ClientTop       =   225
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtClave 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "Clave de la función"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox TxtDescripcion 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Descripcion de la funcion"
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1560
      MouseIcon       =   "FrmFunciones.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Da de alta al nuevo empleado"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4080
      MouseIcon       =   "FrmFunciones.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Regresa a la opcion anterior"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Clave:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion:"
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   885
   End
End
Attribute VB_Name = "FrmFunciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'Las declaraciones de variables seran obligatorias
'Agrega una nueva Funcion
Private Sub CmdAceptar_Click()
Dim SQL As String
On Error GoTo Error
    
    SQL = "Insert Into FUNCIONES Values ('" & TxtClave & "','" & TxtDescripcion & "')"
    Conn.BeginTrans
    Conn.Execute SQL
    Conn.CommitTrans
    IMPRIME "Función Registrada correctamente"
    
    If MsgBox("Desea agregar otra Función", vbQuestion + vbYesNo) = vbYes Then
        TxtClave = ""
        TxtDescripcion = ""
        TxtClave.SetFocus
    Else
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

Sub Errores()
    Select Case Err.Number
'        Case -2147467259        'Llave primaria duplicada
'            IMPRIME "sLa clave del empleado ya esta en uso"
'            TxtClave.SetFocus
        Case Else
            IMPRIME Err.Description
    End Select
    
    Conn.RollbackTrans  'Deshago la transaccion
    Exit Sub
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
    If KeyAscii = 13 And TxtDescripcion <> "" Then CmdAceptar.SetFocus
End Sub
