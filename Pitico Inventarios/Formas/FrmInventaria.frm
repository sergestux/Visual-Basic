VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmInventaria 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inventariar"
   ClientHeight    =   6765
   ClientLeft      =   210
   ClientTop       =   315
   ClientWidth     =   11775
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdTerminarInventario 
      Caption         =   "Terminar Inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      MouseIcon       =   "FrmInventaria.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FrmInventaria.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   0
      TabIndex        =   5
      Top             =   1725
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "PRODUCTOS"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameConteo 
      Caption         =   "Conteo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton CmdTerminarConteo 
         Caption         =   "Terminar Conteo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7920
         MouseIcon       =   "FrmInventaria.frx":0BD4
         MousePointer    =   99  'Custom
         Picture         =   "FrmInventaria.frx":0EDE
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtConteo 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "1"
         ToolTipText     =   "Numero de conteo (1,2 o 3)"
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox TxtUbicacion 
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Text            =   "R01"
         ToolTipText     =   "Donde se encuentra ubicado el producto"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label LblResponsableConteo 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4600
         TabIndex        =   15
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Conteo"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion"
         Height          =   195
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      MouseIcon       =   "FrmInventaria.frx":12AC
      MousePointer    =   99  'Custom
      Picture         =   "FrmInventaria.frx":15B6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdContar 
      Caption         =   "Inventariar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      MouseIcon       =   "FrmInventaria.frx":1694
      MousePointer    =   99  'Custom
      Picture         =   "FrmInventaria.frx":199E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TxtCajas 
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      ToolTipText     =   "Num. de cajas del producto"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox TxtCodigo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Código de barras del producto"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox TxtDescripcion 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      ToolTipText     =   "Descripcion del producto"
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Producto"
      Height          =   195
      Left            =   2520
      TabIndex        =   11
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cajas"
      Height          =   195
      Left            =   8400
      TabIndex        =   8
      Top             =   1320
      Width           =   390
   End
End
Attribute VB_Name = "FrmInventaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit         'Las declaraciones de variables seran obligatorias
Dim BANDERA As Boolean  'Para saber si presione enter en el grid
Dim CONS As String      'Para almacenar el numero consecutivo del producto si es que se tiene
Dim CODIGO As String    'Para almacenar el Codigo del producto si es que se tiene

'Actualiza su nuevo codigo de barras o descripcion
Private Sub CmdActualizar_Click()
Dim SQL As String
On Error GoTo Error
    
    SQL = "Update TFPRODUC Set barraspza=" & TxtCodigo & ", Descripc='" & TxtDescripcion & _
    "' Where Barraspza =" & CODIGO & " And Consec='" & CONS & "'"
    'Debug.Print SQL
    Conn.BeginTrans
    Conn.Execute SQL
    Conn.CommitTrans
    
    CmdActualizar.Visible = False
    CmdTerminarInventario.Visible = True
    CmdContar.Enabled = True
    TxtCajas.Enabled = True
    IMPRIME "Datos del producto actualizados correctamente"
    TxtDescripcion_Change  'Simulo haber presionado buscar por descripcion
    TxtCajas.SetFocus
    Exit Sub
Error:
    Errores
End Sub

Private Sub CmdContar_Click()
Dim SQL As String
On Error GoTo Error
If INVENTARIO = "" Then Exit Sub

    If ValidarEntrada = True Then
        Select Case CONTEO
        Case 1
            SQL = "Insert Into CONTEO Values ('" & TxtCodigo & "','" & CONS & "','" & INVENTARIO & "','" & TxtUbicacion & "'," & TxtCajas & ",0,0)"
        Case 2
            SQL = "Insert Into CONTEO Values ('" & TxtCodigo & "','" & CONS & "','" & INVENTARIO & "','" & TxtUbicacion & "',0," & TxtCajas & ",0)"
        Case 3
            SQL = "Insert Into CONTEO Values ('" & TxtCodigo & "','" & CONS & "','" & INVENTARIO & "','" & TxtUbicacion & "',0,0," & TxtCajas & ")"
        End Select
            
        'Debug.Print SQL
        Conn.BeginTrans
        Conn.Execute SQL
        Conn.CommitTrans
        IMPRIME "Producto: " & TxtCodigo & "  " & TxtDescripcion & " Contabilizado correctamente"
        TxtDescripcion = ""
        TxtCodigo = ""
        TxtCajas = ""
        TxtCodigo.SetFocus
        Exit Sub
Error:
    Conn.RollbackTrans
    Errores
    End If
End Sub

Sub Errores()
    Select Case Err.Number
        'Case -2147467259        'Error de integridad referencial
'            IMPRIME "El producto: " & TxtDescripcion & " y Código: " & TxtCodigo & " No se encuentra aun registrado"
        Case -2147217900
            IMPRIME "Faltan Datos"
        Case Else
            IMPRIME Err.Description
    End Select
    
End Sub

'Regresa Verdadero si es una Entrada valida, falso si faltan datos
Function ValidarEntrada() As Boolean
ValidarEntrada = False
    If TxtCodigo.Text = "" Then
        IMPRIME "Falta el codigo del Producto"
        TxtCodigo.SetFocus
    ElseIf TxtDescripcion.Text = "" Then
        IMPRIME "Falta la descripcion del Producto"
        TxtDescripcion.SetFocus
    ElseIf TxtCajas.Text = "" Then
        IMPRIME "Falta la cantidad de cajas"
        TxtCajas.SetFocus
    Else
        ValidarEntrada = True
    End If
End Function
Private Sub CmdSalir_Click()
    Unload Me
End Sub


'Actualizar los datos del correspondiente al conteo
'Se almacena la fecha y hora actual del sistema como datos de fin del inventario
Private Sub CmdTerminarConteo_Click()
If INVENTARIO = "" Then Exit Sub

    CONTAR  'Realiza las cuentas del conteo actual
    If (CONTEO + 1) <= 3 Then  'Si Conteo= 1, 2 o 3
        If MsgBox("¿Desea comenzar otro conteo?", vbQuestion + vbYesNo) = vbYes Then
            CONTEO = Val(TxtConteo) + 1
            
            Conn.BeginTrans
            Conn.Execute "UPDATE Inventario SET Conteo=" & CONTEO & _
            ", FECHA_FIN='15/05/06', HORA_FIN='12:00'" & _
            " WHERE Clave='" & INVENTARIO & "'"
            Conn.CommitTrans
            
            TxtConteo.Enabled = True
            TxtConteo = CONTEO
            TxtConteo.Enabled = False
            TxtUbicacion.SetFocus
        End If
    Else
        IMPRIME "SOLO TRES CONTEOS POR AHORA"
    End If
End Sub

'Saco los totales del conteo y actualizo la tabla Inventarios
Sub CONTAR()
Dim SQL As String
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
Dim TOTAL As Currency   'ALMACENA TEMPORALMENTE EL TOTAL DEL CONTEO ACTUAL
On Error GoTo Error
    'Saco el total del conteo actual
    SQL = "SELECT SUM(C.CONTEO" & CONTEO & " * P.COSTOCAJ) AS `CONTEO` FROM TFPRODUC P, CONTEO C " & _
    " WHERE P.CONSEC = C.CONSEC AND P.BARRASPZA = C.PRODUCTO AND " & _
    " C.INVENTARIO = '" & INVENTARIO & "'"
    'Debug.Print SQL
    Set TABLA = Conn.Execute(SQL)
    TOTAL = TABLA.Fields("CONTEO").Value
    
    SQL = "UPDATE Inventario SET TOTAL_CONTEO" & CONTEO & "=" & TOTAL
    SQL = SQL & " WHERE Clave='" & INVENTARIO & "'"
    
    'Debug.Print SQL
    Conn.BeginTrans
    Conn.Execute SQL
    Conn.CommitTrans
    Exit Sub
Error:
    Conn.RollbackTrans
    Errores

End Sub
'Hace los procesos correspondientes cuando se termina el inventario
Private Sub CmdTerminarInventario_Click()
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
Dim SQL As String

If INVENTARIO = "" Or CONTEO = 0 Then Exit Sub
    If MsgBox("¿Realmente desea dar por terminado este inventario?", vbYesNo + vbQuestion) = vbYes Then
        CONTAR
        
        SQL = "UPDATE INVENTARIO SET FECHA_FIN='" & Format(Date, "dd/mm/yyyy") & "', HORA_FIN='" & Format(Time, "HH:MM") & "'," & _
        " ESTADO= YES WHERE CLAVE='" & INVENTARIO & "'"  'Actualizo el estado del inventario
        'Debug.Print SQL
        Conn.BeginTrans
        Conn.Execute SQL
        Conn.CommitTrans
        Unload Me
        FrmReportes.Show
    End If
End Sub

Private Sub DataGrid1_DblClick()
    CmdActualizar.Visible = True
    CmdTerminarInventario.Visible = False
    CODIGO = DataGrid1.Columns(0).Value
    CONS = DataGrid1.Columns("CONSEC").Value
    TxtCodigo = DataGrid1.Columns(0).Value
    TxtDescripcion = DataGrid1.Columns(1).Value
    TxtCodigo.SetFocus
    TxtCajas.Enabled = False
    CmdContar.Enabled = False
End Sub

'Se selecciona un  producto para contarlo
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'If se presiona enter
        BANDERA = True
        CODIGO = DataGrid1.Columns(0).Value
        CONS = DataGrid1.Columns("CONSEC").Value
        'CONS = RTrim(DataGrid1.Columns("CONSEC").Value)
        TxtCodigo = DataGrid1.Columns(0).Value
        TxtDescripcion = DataGrid1.Columns(1).Value
        If CmdActualizar.Visible = False Then TxtCajas.SetFocus
        'SendKeys "{TAB}"    'Paso el enfoque al siguiente control
    End If
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    ActualizarGRID DE1.Commands("BuscarProducto").CommandText
    BANDERA = False
    Me.Caption = "Tienda: " & TIENDA & "  Inventario: " & INVENTARIO & " Conteo: " & CONTEO & " Ubicacion: " & UBICACION
    Me.TxtUbicacion = UBICACION
    Me.TxtConteo = CONTEO
    Me.TxtConteo.Enabled = False
    'Me.CmdTerminarConteo = False
End Sub

Private Sub TxtCajas_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
    If KeyAscii = 13 Then
        CmdContar.SetFocus
    End If
End Sub

'Private Sub TxtCajas_GotFocus()
'    SendKeys "+ ({END})"
'End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
Dim SQL As String
    SoloNumeros KeyAscii
    If KeyAscii = 13 Then
        If CmdActualizar.Visible = False Then
            If TxtCodigo <> "" Then   'Si hay al menos algo para buscar
                SQL = "SELECT BARRASPZA AS `Código de Barras`, DESCRIPC AS Producto," & _
                " (paquetes & "" x "" & contenid & "" "" & medida) AS Presentacion, " & _
                " Costocaj AS `Costo x Caja`, CLAPROVE As PROV, CONSEC From TFPRODUC "
                '" CLAPROVE As PROV From TFPRODUC WHERE (BAJA <> '1')"
                SQL = SQL & " WHERE BARRASPZA =" & TxtCodigo
                SQL = SQL & " ORDER BY DESCRIPC, BARRASPZA"
                'Debug.Print SQL
                ActualizarGRID SQL
                If TABLA.State = adStateOpen Then TABLA.Close
                Set TABLA = Conn.Execute(SQL)
                If TABLA.RecordCount > 0 Then
                    DataGrid1.SetFocus
                Else
                    CODIGO = ""
                    IMPRIME "Producto no encontrado"
                End If
            End If
        End If
        Exit Sub
    End If
End Sub

Private Sub TxtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then 'Si se presiona la flecha hacia abajo
        If DataGrid1.Enabled = True Then DataGrid1.SetFocus
    End If

End Sub

Private Sub TxtConteo_Change()
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
        
    Conn.BeginTrans
    Conn.Execute "UPDATE Inventario SET Conteo=" & CONTEO & " WHERE Clave='" & INVENTARIO & "'"
    Conn.CommitTrans
'/*****PARA ENCONTRAR AL RESPONSABLE DEL CONTEO*******/
'    If TABLA.State = adStateOpen Then TABLA.Close
'        Select Case Val(TxtConteo)
'            Case 1
'                Set TABLA = Conn.Execute("SELECT EMPLEADOS.NOMBRE From INVENTARIO, EMPLEADOS WHERE INVENTARIO.RESP_CONTEO1 = EMPLEADOS.CLAVE AND INVENTARIO.CLAVE = '" & INVENTARIO & "'")
'
'            Case 2
'                Set TABLA = Conn.Execute("SELECT EMPLEADOS.NOMBRE From INVENTARIO, EMPLEADOS WHERE INVENTARIO.RESP_CONTEO2 = EMPLEADOS.CLAVE AND INVENTARIO.CLAVE = '" & INVENTARIO & "'")
'
'            Case 3
'                Set TABLA = Conn.Execute("SELECT EMPLEADOS.NOMBRE From INVENTARIO, EMPLEADOS WHERE INVENTARIO.RESP_CONTEO3 = EMPLEADOS.CLAVE AND INVENTARIO.CLAVE = '" & INVENTARIO & "'")
'            Case 0
'                Exit Sub
'        End Select
'
'    If TABLA.RecordCount > 0 Then
'        LblResponsableConteo = TABLA.Fields(0).Value
         TxtConteo.Enabled = False
         If Val(TxtConteo) = 3 Then CmdTerminarConteo.Enabled = False
'    Else
'        LblResponsableConteo = "No se encuentra el responsable del conteo"
'        IMPRIME "Verifique si el Inventario actual ya tiene un responsable de este conteo"
'    End If
'
End Sub

'Private Sub TxtConteo_Validate(Cancel As Boolean)
'    If Val(TxtConteo) < 1 Or Val(TxtConteo) > 3 Then
'        IMPRIME "Escriba 1, 2 o 3, Segun el conteo correspondiente"
'        Cancel = True     'Dejo que Pierda el enfoque
'    Else
'        Cancel = False  'Dejo que Pierda el enfoque
'        TxtUbicacion.SetFocus
'    End If
'End Sub

Private Sub TxtDescripcion_Change()
Dim SQL As String

    If TxtDescripcion.Text <> "" Then   'Si hay al menos algo para buscar
        SQL = "SELECT BARRASPZA AS `Código de Barras`, DESCRIPC AS Producto," & _
            " (paquetes & "" x "" & contenid & "" "" & medida) AS Presentacion, " & _
            " Costocaj AS `Costo x Caja`, CLAPROVE As PROV, CONSEC From TFPRODUC "
        If BANDERA = True Then  'Si la busqueda vino del grid
            SQL = SQL & " WHERE Barraspza=" & CODIGO & " AND CONSEC='" & CONS & "'"
        ElseIf CmdActualizar.Visible = False Then
            SQL = SQL & " WHERE (DESCRIPC LIKE '%" & TxtDescripcion.Text & "%')"
        End If
        
        SQL = SQL & " ORDER BY DESCRIPC, BARRASPZA"
        BANDERA = False         'La busqueda vino de esta caja
        ActualizarGRID SQL
    End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
    
    If KeyAscii = 13 Then
        If CmdActualizar.Visible = True Then
            CmdActualizar.SetFocus
        Else
            CmdContar.SetFocus
        End If
    End If
End Sub

Private Sub TxtDescripcion_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then 'Si se presiona la flecha hacia abajo
        If DataGrid1.Enabled = True Then DataGrid1.SetFocus
    End If
End Sub

'Actualiza el GRID con la consulta que se le pase
Sub ActualizarGRID(SQL As String)
Dim TABLA As New ADODB.Recordset    'La tabla donde se haran los trabajos
    'Debug.Print SQL
    If TABLA.State = adStateOpen Then TABLA.Close
    Set TABLA = Conn.Execute(SQL)
    Set DataGrid1.DataSource = TABLA
    'DataGrid1.Refresh
    DataGrid1.Columns.Item(0).Width = 102        'Código de Barras
    DataGrid1.Columns.Item(1).Width = 429        'Producto
    DataGrid1.Columns.Item(2).Width = 87         'Presentacion
    
    DataGrid1.Columns.Item(3).Width = 79         'Costo x Caja
    DataGrid1.Columns.Item(3).NumberFormat = "$  ###,###,###.00"     'Costo x Caja
    DataGrid1.Columns.Item(3).Alignment = dbgRight
    
    DataGrid1.Columns.Item(4).Width = 50         'Proveedor
    DataGrid1.Columns.Item(4).Alignment = dbgRight
    
    DataGrid1.Columns.Item(5).Visible = False    'BAJA
    DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
    DataGrid1.Refresh
    
    If TABLA.RecordCount > 0 Then
        DataGrid1.Enabled = True
    Else
        DataGrid1.Enabled = False
    End If
End Sub

Private Sub TxtUbicacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierto a mayuculas lo tecleado
    If KeyAscii = 13 Then
        If TxtUbicacion <> "" Then
            Conn.Execute "UPDATE Inventario SET Ubicacion='" & TxtUbicacion & "'"
            TxtCodigo.SetFocus
            UBICACION = TxtUbicacion
        End If
    End If
End Sub
