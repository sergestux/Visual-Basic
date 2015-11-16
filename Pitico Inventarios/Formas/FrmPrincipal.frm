VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FrmPrincipal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "HOLDING MÉXICO CENTRO AMERICA S.A. DE C.V"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11880
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "FrmPrincipal.frx":0442
   ScrollBars      =   0   'False
   Begin MSComctlLib.StatusBar Barra 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7845
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15293
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "13/09/06"
            Object.ToolTipText     =   "Fecha Actual"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11:45"
            Object.ToolTipText     =   "Hora Actual"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":322F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":32613
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":36123
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":36575
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":36FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":372F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   2487
      ButtonHeight    =   1005
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Empleados"
            Key             =   "Empleados"
            Object.ToolTipText     =   "Empleados"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tiendas"
            Key             =   "Tiendas"
            Object.ToolTipText     =   "Datos de la tienda inventariada"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Inventario"
            Key             =   "Inventario"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contar"
            Key             =   "Contar"
            Object.ToolTipText     =   "Comenzar a contar los productos"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reportes"
            Key             =   "Reportes"
            Object.ToolTipText     =   "Reportes del inventario"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Terminar el programa"
            ImageIndex      =   6
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "FrmPrincipal.frx":37748
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'Las declaraciones de variables seran obligatorias

Private Sub MDIForm_Load()
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
Dim SQL As String
    
    If TABLA.State = adStateOpen Then TABLA.Close
    Set TABLA = Conn.Execute("SELECT CLAVE, TIENDA AS TDA, RESP_TIENDA, RESP_INVENTARIO, RESP_CONTEO1, RESP_CONTEO2, RESP_CONTEO3, CONTEO AS CNT, UBICACION FROM INVENTARIO WHERE ESTADO=NO")
    
    If TABLA.RecordCount > 0 Then INVENTARIO = TABLA.Fields("CLAVE").Value            'Recupero el inventario si lo hay
    
    'Toolbar1.Buttons.Item("Reportes").Enabled = False
    'Toolbar1.Buttons.Item("Salir").Enabled = True
    
    If INVENTARIO <> "" Then    'Si existe un inventario Sin terminar
        If MsgBox("Desea Continuar con el inventario: '" & INVENTARIO & "'?", vbQuestion + vbYesNo) = vbYes Then
            
            CONTEO = TABLA!CNT       'Recupero el Conteo del Inventario si lo hay
            TIENDA = TABLA!TDA   'Actualizo la Variable de tiendas
            UBICACION = TABLA!UBICACION
            Toolbar1.Buttons.Item("Empleados").Enabled = False
            Toolbar1.Buttons.Item("Tiendas").Enabled = False

            FrmInventario.Show
        Else
            FrmEmpleados.Show
        End If
    
    Else
        Toolbar1.Buttons.Item("Contar").Enabled = False
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("¿Realmente desea salir del sistema?", vbQuestion + vbYesNo) = vbNo Then Cancel = 1
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
    
    'PONGO LA VARIABLE DE ERROR A FALSO
    Conn.BeginTrans
    Conn.Execute "UPDATE CONTROL SET ERROR= NO"
    Conn.CommitTrans
    
    'Cierro la conexion a la Base de datos
    If Conn.State = adStateOpen Then Conn.Close
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button
        Case "Tiendas"
            FrmTienda.Show
        
        Case "Empleados"
            FrmEmpleados.Show
            
        Case "Inventario"
            FrmInventario.Show
        
        Case "Contar"
            FrmInventaria.Show
        
        Case "Reportes"
            FrmReportes.Show
            
        Case "Salir"
            Unload Me
        
    End Select
End Sub
