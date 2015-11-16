VERSION 5.00
Begin VB.Form fcod 
   Caption         =   "Impresion de Codigos de Barras"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "CODIGOS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtpreaut 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkPrecio 
         Caption         =   "Incluir precio de autoservicio"
         Height          =   495
         Left            =   3840
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tamaño 2"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tamaño 1 "
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   11
         Top             =   960
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.TextBox txtpaq 
         Height          =   285
         Left            =   5400
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtcodigo 
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtdos 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   6255
      End
      Begin VB.TextBox txtuno 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2760
         Width           =   6255
      End
      Begin VB.CheckBox chkpaq 
         Caption         =   "Incluir Paquetes"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtcopias 
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Text            =   "1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   4560
         Picture         =   "CODIGOS.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Texto a Incluir en Etiqueta"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo de Barras"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "No. de Copias :"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "fcod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not verImpresora Then Exit Sub
If Option1.Item(1).Value Then
   tam = 1
Else
   tam = 2
End If
If Len(txtcodigo.Text) > 7 Then
   MsgBox "No se permiten Codigos Mayores  a 7 digitos ", vbInformation, "CODIGOS"
   Exit Sub
End If
If Len(txtcodigo.Text) < 2 Then
   MsgBox "No se permiten Codigos Menores a  1 digitos ", vbInformation, "CODIGOS"
   Exit Sub
End If
nAncho = 250    'En puntos
Printer.ScaleMode = vbPoints
Printer.CurrentX = 10
espacio = "   "
For i = 1 To Val(txtcopias.Text)
    Printer.Font = "arial"
    Printer.FontSize = "8"
    Printer.Print espacio & Trim(txtuno.Text)
    PRECIO = Space(5) & IIf(chkPrecio.Value = 1, txtpreaut.Text, "")
    If chkpaq.Value = 1 Then
       Printer.Print Trim(txtpaq.Text) & " x " & Trim(txtdos.Text) & PRECIO
    Else
       Printer.Print Space(10) & Trim(txtdos.Text) & PRECIO
    End If
    If tam = 1 Then
        Printer.Font = "ZB 39* 15mil/2:1"
    ElseIf tam = 2 Then
        Printer.Font = "ZB 39* 10mil/2:1"
    End If
    Printer.FontSize = 40
    'Printer.FontSize = 30
    Printer.CurrentX = 10
    Printer.Print (txtcodigo.Text)
    Printer.EndDoc
Next
End Sub

Private Sub Form_Load()
txtcodigo.Text = fhojacat.Adoprod.Recordset!barraspza
txtuno.Text = fhojacat.Adoprod.Recordset!DESCRIPC
txtdos.Text = fhojacat.Adoprod.Recordset!Contenid & " x " & fhojacat.Adoprod.Recordset!Medida
txtpaq.Text = fhojacat.Adoprod.Recordset!PAQUETES
txtpreaut.Text = Format(fhojacat.Adoprod.Recordset!precio1, "$###,###,###.00")
End Sub
