VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmServidor 
   Caption         =   "FrmServidor"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Text            =   "ADIOS;SYS_DELEON;RF;CON"
      Top             =   3000
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   120
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label IpLocal 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
   Begin VB.Label LblPuertoLocal 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "FrmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const IPServidor = "148.222.140.198"    'IP DEL SERVIDOR
Const PuertoCliente = 1187     'Puerto por el que se comunican los clientes
Const PuertoLocal = 1186        'Puerto por el que escucha el servidor
Public CONN As New ADODB.Connection
Public RS As New ADODB.Recordset

Private Sub Form_Load()
'    Socket.Bind PuertoLocal  'Puerto por donde escuchara
'    Socket.RemotePort = PuertoCliente   'Puerto por donde escucha el cliente
'    LblPuertoLocal = "Puerto local: " & Socket.LocalPort
'    IpLocal = "IP Local: " & Socket.LocalIP
    'CONN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=" & App.Path & "\MENSAJES-SIIA.mdb;Persist Security Info=True"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Socket.State <> sckClosed Then Socket.Close
End Sub

'Cuando llega un mensaje del cliente
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim MENSAJE As String
    Dim TipoMSG As String
    Dim SQL As String
    Dim POS1 As Byte
    Dim POS2 As Byte
    Dim CONT As Byte
    Dim MENSAJES() As String
    
    POS1 = 1
    POS2 = 1
    CONT = 1
    
    Socket.GetData MENSAJE
    
    Do
        POS2 = InStr(POS1, MENSAJE, ";")
        If POS2 = 0 Then Exit Do
        ReDim Preserve MENSAJES(1 To CONT)
        MENSAJES(CONT) = Mid(MENSAJE, POS1, POS2 - POS1)
        POS1 = POS2 + 1
        CONT = CONT + 1
    Loop
    
    TipoMSG = MENSAJES(1)
    
    Select Case TipoMSG
        Case "HOLA"  'Cuando el usuario se conecta al sistema
            
            SQL = "SELECT * FROM  USUARIOS WHERE USUARIO=" & Mid(MENSAJE, 5)
            Set RS = CONN.Execute(SQL)
            If RS.BOF = True Or RS.EOF = True Then  'SI EL USUARIO NO ESTA EN LA BASE DE DATOS
                SQL = "SELECT * FROM  USUARIOS WHERE USUARIO=" & Mid(MENSAJE, 5)        'SE INSERTA EN LA TABLA
            Else
                SQL = "SELECT * FROM  USUARIOS WHERE USUARIO=" & Mid(MENSAJE, 5)        'SE ACTUALIZA SU DIRECCION IP
            End If
            Socket.SendData "BIENVENIDO"
            
        Case "Mensaje"     'Cuando el mensaje es para otro usuario
            
        Case "ADIOS"     'Cuando el usuario se desconecta
        
    End Select
    
    Socket.SendData "Mensaje recibido: " & MENSAJE
End Sub


Private Sub Text1_DblClick()

    Dim MENSAJE As String
    Dim TipoMSG As String
    Dim SQL As String
    Dim POS1 As Byte
    Dim POS2 As Byte
    Dim CONT As Byte
    Dim MENSAJES() As String
    
    
    POS1 = 1
    POS2 = 1
    CONT = 1
    MENSAJE = Text1
    
    Do
        POS2 = InStr(POS1, MENSAJE, ";")
        If POS2 = 0 Then Exit Do
        ReDim Preserve MENSAJES(1 To CONT)
        MENSAJES(CONT) = Mid(MENSAJE, POS1, POS2 - POS1)
        POS1 = POS2 + 1
        CONT = CONT + 1
    Loop
    
    TipoMSG = MENSAJES(1)
    
    CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=C:\Documents and Settings\User\Mis documentos\SIIA\REPORTES\Winsock\SERVIDOR\MENSAJES-SIIA.mdb;Persist Security Info=True"
    
    
    CONN.Database.Connection
    DB.Name
    Select Case TipoMSG
        Case "HOLA"  'Cuando el usuario se conecta al sistema
            
            SQL = "SELECT * FROM  USUARIOS WHERE USUARIO=" & Mid(MENSAJE, 5)
            Set RS = CONN.Execute(SQL)
            If RS.BOF = True Or RS.EOF = True Then  'SI EL USUARIO NO ESTA EN LA BASE DE DATOS
                SQL = "SELECT * FROM  USUARIOS WHERE USUARIO=" & Mid(MENSAJE, 5)        'SE INSERTA EN LA TABLA
            Else
                SQL = "SELECT * FROM  USUARIOS WHERE USUARIO=" & Mid(MENSAJE, 5)        'SE ACTUALIZA SU DIRECCION IP
            End If
            Socket.SendData "BIENVENIDO"
            
        Case "Mensaje"     'Cuando el mensaje es para otro usuario
            
        Case "ADIOS"     'Cuando el usuario se desconecta
        
    End Select
    CONN.Close
End Sub
