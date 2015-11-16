VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmEmpleados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Empleados"
   ClientHeight    =   6315
   ClientLeft      =   645
   ClientTop       =   315
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdNvoEmpleado 
      Caption         =   "&Nuevo Empleado"
      Height          =   495
      Left            =   3360
      MouseIcon       =   "FrmEmpleados.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Agrega un nuevo empleado"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1200
      MouseIcon       =   "FrmEmpleados.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Asignar la funcion elegida al empleado"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   5520
      MouseIcon       =   "FrmEmpleados.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Regresa al menu principal"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame FrameResponsable 
      Caption         =   "Asignacion de Funciones"
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
      TabIndex        =   0
      Top             =   3600
      Width           =   7215
      Begin VB.TextBox TxtNombre 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         ToolTipText     =   "Nombre del empleado"
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox TxtClave 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         ToolTipText     =   "Clave del empleado"
         Top             =   360
         Width           =   615
      End
      Begin MSDataListLib.DataCombo CmbFuncion 
         Bindings        =   "FrmEmpleados.frx":091E
         DataField       =   "CLAVE"
         DataMember      =   "Funciones"
         DataSource      =   "DE1"
         Height          =   315
         Left            =   960
         TabIndex        =   5
         ToolTipText     =   "Funcion del empleado"
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "FUNCION"
         BoundColumn     =   "CLAVE"
         Text            =   "CmbFuncion"
         Object.DataMember      =   "Funciones"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   12
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Funcion:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   -1440
         TabIndex        =   9
         Top             =   1080
         Width           =   480
      End
   End
   Begin MSDataListLib.DataCombo CmbTiendas 
      Bindings        =   "FrmEmpleados.frx":093A
      DataField       =   "CLAVE"
      DataMember      =   "Tiendas"
      DataSource      =   "DE1"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Tienda"
      Top             =   240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "DESCRIPCION"
      BoundColumn     =   "CLAVE"
      Text            =   "CmbTiendas"
      Object.DataMember      =   "Tiendas"
   End
   Begin MSDataGridLib.DataGrid GridEmpleados 
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Tag             =   "0"
      ToolTipText     =   "Empleados de la tienda elegida"
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "EMPLEADOS"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   104.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   104.882
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'Las declaraciones de variables seran obligatorias

Private Sub CmbFuncion_KeyPress(KeyAscii As Integer)
    CmdAceptar.SetFocus
End Sub

Private Sub CmdAceptar_Click()
Dim SQL As String
On Error GoTo Error
    
    SQL = "Insert Into RESPONSABLES Values ('" & TxtClave & "','" & CmbFuncion.BoundText & "')"
    Conn.BeginTrans
    Conn.Execute SQL
    IMPRIME "Responsable Registrado correctamente"
    Conn.CommitTrans
    
    If MsgBox("Desea asignar una funcion a otro empleado", vbQuestion + vbYesNo) = vbYes Then
        TxtClave = ""
        TxtNombre = ""
        TxtClave.SetFocus
    Else
        TIENDA = CmbTiendas.BoundText   'Actualizo la Variable de tiendas
        FrmPrincipal.Toolbar1.Buttons.Item("Inventario").Enabled = True
        FrmInventario.Show
        Unload Me
    End If
    Exit Sub
Error:
    Errores
End Sub

Private Sub CmdNvoEmpleado_Click()
    FrmNvoEmpleado.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
'Segun la tienda que se elige en el combo de tiendas hago el fitro en el grid de empleados
'de acuerdo al personal que labore en la tienda elegida
Private Sub CmbTiendas_Change()
Dim SQL As String
    SQL = "SELECT E.CLAVE, E.NOMBRE FROM EMPLEADOS E, TIENDAS T WHERE E.TIENDA = T.CLAVE AND" & _
    " T.CLAVE = '" & CmbTiendas.BoundText & "'"
    
    If TxtClave <> "" Then SQL = SQL & " And E.Clave=" & Val(TxtClave)
    
    SQL = SQL & " ORDER BY E.NOMBRE"
    ActualizarGRID SQL
End Sub

'Al presionar Enter sobre el combo de tiendas paso el foco al grid de empleados
Private Sub CmbTiendas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then GridEmpleados.SetFocus
End Sub

'Al presionar enter en el grid de empleados actualizo la clave y el
'empleado y paso el foco al combo de funciones
Private Sub GridEmpleados_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If GridEmpleados.Enabled = True Then
            TxtClave = GridEmpleados.Columns("CLAVE").Value
            TxtNombre = GridEmpleados.Columns("NOMBRE").Value
            CmbFuncion.SetFocus
        End If
    End If
End Sub
'Al cargar el formulario desactivo los botones correspondientes
Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    If INVENTARIO = "" Then
        FrmPrincipal.Toolbar1.Buttons.Item("Inventario").Enabled = False
        FrmPrincipal.Toolbar1.Buttons.Item("Contar").Enabled = False
    Else
        FrmPrincipal.Toolbar1.Buttons.Item("Inventario").Enabled = True
        FrmPrincipal.Toolbar1.Buttons.Item("Contar").Enabled = True
    End If
End Sub

'Al presionar enter, Busco al empleado
Private Sub TxtClave_KeyPress(KeyAscii As Integer)
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
Dim SQL As String
    SoloNumeros KeyAscii
    If KeyAscii = 13 And TxtClave <> "" Then
        SQL = "SELECT Nombre, Tienda From EMPLEADOS WHERE Clave=" & Val(TxtClave) & ""
        'Debug.Print SQL
        If TABLA.State = adStateOpen Then TABLA.Close
        Set TABLA = Conn.Execute(SQL)   'Buscar al empleado con la clave digitada
        If TABLA.RecordCount > 0 Then
            CmbTiendas.BoundText = TABLA.Fields("TIENDA").Value
            CmbTiendas_Change
            TxtNombre = TABLA.Fields("NOMBRE").Value
            CmbFuncion.SetFocus
        Else
            If MsgBox("El empleado no se encuentra registrado ¿Desea Registrarlo?", vbQuestion + vbYesNo) = vbYes Then
                FrmNvoEmpleado.Show
                FrmNvoEmpleado.TxtClave = TxtClave
            End If
        End If
    End If
End Sub
'Al presionar enter paso el enfoque al Combo de Funciones
Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
    If KeyAscii = 13 Then CmbFuncion.SetFocus
End Sub

'Funcion para manejar errores
Sub Errores()
    Select Case Err.Number
        Case -2147467259        'Llave primaria duplicada
            IMPRIME "La clave del empleado ya esta en uso"
            TxtClave.SetFocus
        Case Else
            IMPRIME Err.Description
    End Select
    
    Conn.RollbackTrans  'Deshago la transaccion
    Exit Sub
End Sub

'Actualiza el GRID con la consulta que se le pase
Sub ActualizarGRID(SQL As String)
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
    
    'Debug.Print SQL
    If TABLA.State = adStateOpen Then TABLA.Close
    Set TABLA = Conn.Execute(SQL)
    Set GridEmpleados.DataSource = TABLA
    
    GridEmpleados.Columns.Item(0).Width = 840         'Código de Barras
    GridEmpleados.Columns.Item(1).Width = 5805        'Producto
    GridEmpleados.MarqueeStyle = dbgHighlightRowRaiseCell   'Para que se marque todo el renglon
        
    GridEmpleados.Refresh
    
    If TABLA.RecordCount > 0 Then
        GridEmpleados.Enabled = True
    Else
        GridEmpleados.Enabled = False
    End If
End Sub


