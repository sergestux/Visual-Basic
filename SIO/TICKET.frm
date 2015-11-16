VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Printer.Print "VIVERES Y LICORES S.A. DE C.V."
Printer.Print "CARBONERA 1016 COL TRINIDAD DE LAS H."
Printer.Print "OAXACA, OAX. " & Date
Printer.Print Chr(27) + Chr(64) + Chr(27) + Chr(115)
Printer.EndDoc
End Sub
