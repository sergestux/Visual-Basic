VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmCliente 
   Caption         =   "FrmCliente"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtMensaje 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton CmdEnviarMSG 
      Caption         =   "Enviar Mensaje"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   1186
   End
End
Attribute VB_Name = "FrmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const IPServidor = "148.222.140.198"
Const PuertoServidor = 1186
Const PuertoLocal = 1187

Private Sub CmdEnviarMSG_Click()
    Socket.SendData TxtMensaje
End Sub

Private Sub Form_Load()
    With Socket
        .RemoteHost = IPServidor
        .RemotePort = PuertoServidor
        .Bind PuertoLocal
    End With
End Sub

'Recibe los mensajes del servidor
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim MENSAJE As String
    Socket.GetData MENSAJE
    MsgBox MENSAJE
'    FrmMensaje.LblMensaje = MENSAJE
'    FrmMensaje.Show vbModal
End Sub
