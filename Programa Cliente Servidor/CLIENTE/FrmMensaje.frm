VERSION 5.00
Begin VB.Form FrmMensaje 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FrmMensaje"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMensaje.frx":0000
   ScaleHeight     =   1635
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LblMensaje 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
End Sub
