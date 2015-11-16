VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmInventario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inventario"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   2175
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   8295
      Begin VB.TextBox TxtInventario 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "Clave del Inventario"
         Top             =   240
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo CmbTienda 
         Bindings        =   "FrmInventario.frx":0000
         Height          =   315
         Left            =   3720
         TabIndex        =   1
         ToolTipText     =   "Tienda a Inventariar"
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inventario:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tienda:"
         Height          =   195
         Left            =   2880
         TabIndex        =   17
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Responsables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   8295
      Begin VB.CheckBox Check 
         Caption         =   "Check"
         Height          =   255
         Index           =   4
         Left            =   7800
         TabIndex        =   23
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check 
         Caption         =   "Check"
         Height          =   255
         Index           =   3
         Left            =   7800
         TabIndex        =   22
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check 
         Caption         =   "Check"
         Height          =   255
         Index           =   2
         Left            =   7800
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check 
         Caption         =   "Check"
         Height          =   255
         Index           =   1
         Left            =   7800
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check 
         Caption         =   "Check"
         Height          =   255
         Index           =   0
         Left            =   7800
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSDataListLib.DataCombo CmbResponsable 
         Bindings        =   "FrmInventario.frx":000B
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   2
         ToolTipText     =   "Responsable de la Tienda (Gerente)"
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo CmbResponsable 
         Bindings        =   "FrmInventario.frx":0016
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Responsable del Inventario (Auditor)"
         Top             =   1080
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo CmbResponsable 
         Bindings        =   "FrmInventario.frx":0021
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   4
         ToolTipText     =   "La persona que hace el conteo"
         Top             =   1560
         Visible         =   0   'False
         Width           =   4930
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo CmbResponsable 
         Bindings        =   "FrmInventario.frx":002C
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   5
         ToolTipText     =   "La persona que hace el conteo"
         Top             =   2160
         Visible         =   0   'False
         Width           =   4930
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo CmbResponsable 
         Bindings        =   "FrmInventario.frx":0037
         Height          =   315
         Index           =   4
         Left            =   2520
         TabIndex        =   6
         ToolTipText     =   "La persona que hace el conteo"
         Top             =   2760
         Visible         =   0   'False
         Width           =   4930
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable del 3° Conteo:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable del 2° Conteo:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable del 1° Conteo:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable del Inventario:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1980
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsable de Tienda:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1740
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BORRAR TABLAS"
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1680
      MouseIcon       =   "FrmInventario.frx":0042
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Registra el inventario"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4320
      MouseIcon       =   "FrmInventario.frx":034C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Regresa al menu principal"
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'Las declaraciones de variables seran obligatorias

Private Sub CmbResponsable_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index < 1 Then
        CmbResponsable(Index + 1).SetFocus
    Else
        CmdAceptar.SetFocus
    End If
    
    
    
'    If Index < 4 Then
'        CmbResponsable(Index + 1).SetFocus
'    Else
'        CmdAceptar.SetFocus
'    End If
End Sub
'Al
Private Sub CmbTienda_Change()
    ActualizaComboRe 0, "'RESTDA'"      'Actualizo el combo de Responsable de Tienda
    ActualizaComboRe 1, "'RESINV'"      'Actualizo el combo de Responsable de Inventarios
    ActualizaComboRe 2, "'RESCNT'"      'Actualizo el combo de Responsable del 1° Conteo
    ActualizaComboRe 3, "'RESCNT'"      'Actualizo el combo de Responsable del 2° Conteo
    ActualizaComboRe 4, "'RESCNT'"      'Actualizo el combo de Responsable del 3° Conteo
End Sub

Private Sub CmbTienda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmbResponsable(0).SetFocus
End Sub

'Private Sub CmbTienda_Click(Area As Integer)
'    ActualizaComboRe 0, "'RESTDA'"      'Actualizo el combo de Responsable de Tienda
'    ActualizaComboRe 1, "'RESINV'"      'Actualizo el combo de Responsable de Inventarios
'    ActualizaComboRe 2, "'RESCNT'"      'Actualizo el combo de Responsable del 1° Conteo
'    ActualizaComboRe 3, "'RESCNT'"      'Actualizo el combo de Responsable del 2° Conteo
'    ActualizaComboRe 4, "'RESCNT'"      'Actualizo el combo de Responsable del 3° Conteo
'End Sub

Private Sub CmdAceptar_Click()
Dim SQL As String
On Error GoTo Error
  
    INVENTARIO = TxtInventario
    TIENDA = CmbTienda.BoundText
    
    If INVENTARIO = "" Then     'Si no hay ningun inventario cargado aun
        INVENTARIO = TIENDA & Format(Date, "ddmmyy")    'Genero la clave del Inventario
        If MsgBox("Se genero la clave de Inventario: '" & INVENTARIO & "' Desea usar esa clave", vbQuestion + vbYesNo) = vbNo Then
            INVENTARIO = ""
            TxtInventario.SetFocus
            Exit Sub
        End If
    End If
    
    '****************ACTUALIZO EL INVENTARIO ACTUAL**********************
    SQL = "Insert Into INVENTARIO Values ('" & INVENTARIO & "','" & Format(Date, "dd/mm/yyyy") & "','" & Format(Time, "HH:MM") & "','" & Format(Date, "dd/mm/yyyy") & "','" & Format(Time, "HH:MM") & "','" & _
    CmbTienda.BoundText & "'," & _
    CmbResponsable(0).BoundText & "," & CmbResponsable(1).BoundText & "," & IIf(CmbResponsable(2) = "", 0, CmbResponsable(2).BoundText) & "," & IIf(CmbResponsable(3) = "", 0, CmbResponsable(3).BoundText) & "," & IIf(CmbResponsable(4) = "", 0, CmbResponsable(4).BoundText) & ",0,0,0,1,'',No)"
    'Debug.Print SQL
    Conn.BeginTrans
    Conn.Execute SQL
    Conn.CommitTrans
    
    CONTEO = 1      'Es el primer conteo
    TxtInventario = INVENTARIO
    IMPRIME "Los datos del inventario se guardaron correctamente"
    FrmPrincipal.Toolbar1.Buttons.Item("Contar").Enabled = True
    
    MsgBox "Inventario registrado correctamente", vbExclamation
    
    FrmInventaria.Show
    Unload Me
    
    Exit Sub            'Salgo del procedimiento
Error:
    Conn.RollbackTrans      'SI hay un error se deshace la transacion
    Errores
    
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Sub Errores()
INVENTARIO = ""
    Select Case Err.Number
        Case -2147467259        'Llave primaria duplicada
            IMPRIME "Esta Clave de Inventario ya esta en uso"
        Case -2147217900
            IMPRIME "Faltan Datos"
        Case Else
            IMPRIME Err.Description
    End Select
    
End Sub

Private Sub Command2_Click()
    Conn.Execute "delete from inventario"
    Conn.Execute "delete from responsables"
    IMPRIME "Hecho"
    TIENDA = ""
    INVENTARIO = ""
    CONTEO = 0
    UBICACION = ""
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim SQL As String
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos

    Me.Left = 0
    Me.Top = 0
    'Si Hay un inventario en proceso lo cargo en el Formulario
    If INVENTARIO <> "" Then
        If TABLA.State = adStateOpen Then TABLA.Close
        Set TABLA = Conn.Execute("SELECT RESP_TIENDA, RESP_INVENTARIO, RESP_CONTEO1, RESP_CONTEO2, RESP_CONTEO3 FROM INVENTARIO WHERE Clave='" & INVENTARIO & "'")
        Me.CmbTienda.BoundText = TIENDA                  'Actualizo el combo de tiendas
        Me.CmbTienda.Enabled = False
        Me.CmbResponsable(0).BoundText = TABLA.Fields("RESP_TIENDA").Value     'Actualizo el combo de tiendas
        Me.CmbResponsable(0).Enabled = False
        Me.CmbResponsable(1).BoundText = TABLA.Fields("RESP_INVENTARIO").Value 'Actualizo el combo de tiendas
        Me.CmbResponsable(1).Enabled = False
        Me.CmbResponsable(2).BoundText = TABLA.Fields("RESP_CONTEO1").Value    'Actualizo el combo de tiendas
        Me.CmbResponsable(2).Enabled = False
        Me.CmbResponsable(3).BoundText = TABLA.Fields("RESP_CONTEO2").Value    'Actualizo el combo de tiendas
        Me.CmbResponsable(3).Enabled = False
        Me.CmbResponsable(4).BoundText = TABLA.Fields("RESP_CONTEO3").Value    'Actualizo el combo de tiendas
        Me.CmbResponsable(4).Enabled = False
        Me.CmdAceptar.Enabled = False    'Desactivo el boton de aceptar
        Me.TxtInventario = INVENTARIO
        Me.TxtInventario.Enabled = False
        FrmPrincipal.Toolbar1.Buttons.Item("Contar").Enabled = True
    Else

        'Toolbar1.Buttons.Item("Inventario").Enabled = False
        FrmPrincipal.Toolbar1.Buttons.Item("Contar").Enabled = False

'        Toolbar1.Buttons.Item("Inventario").Enabled = True
'        Toolbar1.Buttons.Item("Contar").Enabled = True


        
        '/*******Cargo el catalogo de tiendas en el Combo CmbTiendas***************/
        SQL = "SELECT CLAVE, CLAVE & "" "" & DESCRIPCION AS DESCRIPCION From TIENDAS ORDER BY 2"
        If TABLA.State = adStateOpen Then TABLA.Close
        Set TABLA = Conn.Execute(SQL)
        Set CmbTienda.DataSource = TABLA
        Set CmbTienda.RowSource = TABLA
        CmbTienda.DataField = "CLAVE"
        CmbTienda.BoundColumn = "CLAVE"
        CmbTienda.ListField = "DESCRIPCION"
        CmbTienda.Refresh
        
        If TIENDA <> "" Then CmbTienda.BoundText = TIENDA
        
    End If
End Sub

Sub ActualizaComboRe(NumCombo As Integer, Funcion As String)
Dim SQL As String
Dim RESPONSABLES As New ADODB.Recordset
    'SQL = "SELECT E.CLAVE, E.NOMBRE FROM EMPLEADOS E, RESPONSABLES R, FUNCIONES F WHERE E.CLAVE = R.CLAVE AND R.FUNCION = F.CLAVE "
    SQL = "SELECT E.CLAVE, E.CLAVE & "" "" & E.NOMBRE AS NOMBRE FROM EMPLEADOS E, RESPONSABLES R, FUNCIONES F WHERE E.CLAVE = R.CLAVE AND R.FUNCION = F.CLAVE " & _
    " AND F.CLAVE = " & Funcion
    
    TIENDA = CmbTienda.BoundText
    
    If TIENDA <> "" Then
        'El responsable del inventario trabaja en Oficinas
        If NumCombo = 1 Then
            SQL = SQL & " AND E.TIENDA = '023'"
        Else
            SQL = SQL & " AND E.TIENDA = '" & TIENDA & "'"
        End If
        
    End If

    SQL = SQL & " ORDER BY E.NOMBRE"
   
    'Debug.Print SQL
    If RESPONSABLES.State = adStateOpen Then RESPONSABLES.Close
    
    Set RESPONSABLES = Conn.Execute(SQL)
    
    'If RESPONSABLES.RecordCount > 0 Then
        Set CmbResponsable(NumCombo).DataSource = RESPONSABLES
        Set CmbResponsable(NumCombo).RowSource = RESPONSABLES
        CmbResponsable(NumCombo).DataField = "CLAVE"
        CmbResponsable(NumCombo).BoundColumn = "CLAVE"
        CmbResponsable(NumCombo).ListField = "NOMBRE"
        CmbResponsable(NumCombo).Refresh
    'End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If INVENTARIO <> "" Then FrmInventaria.Show
End Sub

Private Sub TxtInventario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
End Sub
