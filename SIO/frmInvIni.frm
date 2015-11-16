VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvIni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura de Inventario Inicial"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmInvIni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar STB1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   3255
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   "                          Para salir presione la tecla Escape"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCajas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtClave 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label LblEtiquetas 
      Caption         =   "Cajas"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label LblEtiquetas 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label LblEtiquetas 
      Caption         =   "Clave del producto"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmInvIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtCajas_GotFocus()
  SendKeys "+{END}"
End Sub

Private Sub txtCajas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub txtCajas_LostFocus()
If MsgBox("ESTAN CORRECTOS LOS DATOS?", vbQuestion + vbYesNo) = vbYes Then
   'cn.Execute "UPDATE INVENTARIO SET InInicial = " & CStr(txtCajas.Text) & ", InCant = " & CStr(txtCajas.Text) & "* Paquetes from tfproduc WHERE Inprod = '" & Trim(txtClave.Text) & "' AND Inprod = consec"
   cn.Execute "UPDATE INVENTARIO SET InCant = " & CStr(txtCajas.Text) & " from tfproduc WHERE Inprod = '" & Trim(txtClave.Text) & "' AND Inprod = consec"
   txtClave.Text = ""
   txtCajas.Text = ""
   LblEtiquetas(1).Caption = ""
Else
   'txtCajas.SetFocus
End If
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
Dim rsttemp As ADODB.Recordset
If KeyAscii = 13 Then

  Set rsttemp = New ADODB.Recordset
  rsttemp.Open "SELECT descripc,LTrim(str(Paquetes)) + ' X ' + LTRIM( str(contenid,10,3)) + ' ' + MEDIDA as PRESENT, INPROD, InInicial, InInicialp,INCANT FROM TFPRODUC,INVENTARIO WHERE consec ='" & txtClave.Text & "' AND Inprod = consec", cn, adOpenKeyset, adLockOptimistic, adCmdText
  If rsttemp.RecordCount > 0 Then
    Printer.ScaleMode = vbPoints
    LblEtiquetas(1).Caption = rsttemp!descripc & Chr(13) & rsttemp!Present
    txtCajas.Text = rsttemp!IncANT
    SendKeys vbTab
    KeyAscii = 0
  Else
    MsgBox "NO EXISTE EN EL INVENTARIO EL PRODUCTO CON LA CLAVE ESPECIFICADA", vbExclamation
    txtClave.SetFocus
  End If

ElseIf KeyAscii = 27 Then
   Unload Me
End If
End Sub


Private Sub txtPiezas_GotFocus()
  SendKeys "+{END}"
End Sub

Private Sub txtPiezas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub

Private Sub txtUbi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub
