VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1590
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   939.424
   ScaleMode       =   0  'User
   ScaleWidth      =   4422.435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2130
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2130
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre de usuario:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1380
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Contrase�a:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'Las declaraciones de variables seran obligatorias

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesi�n fallido
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'comprobar si la contrase�a es correcta
    If txtPassword = "password" Then
        'colocar c�digo aqu� para pasar al sub
        'que llama si la contrase�a es correcta
        'lo m�s f�cil es establecer una variable global
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "La contrase�a no es v�lida. Vuelva a intentarlo", , "Inicio de sesi�n"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
