VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmFactDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de la facturas"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frmfactdet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcampos 
      DataField       =   "concepto"
      DataSource      =   "AdoFacturas"
      Height          =   615
      Index           =   8
      Left            =   1080
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   450
      Left            =   2640
      Picture         =   "frmfactdet.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Restaura la factura cancelada"
      Top             =   0
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   6360
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoFacturas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btndes 
      Caption         =   "&Descancelar"
      Height          =   450
      Left            =   3840
      Picture         =   "frmfactdet.frx":05B4
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Restaura la factura cancelada"
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame fraCon 
      BackColor       =   &H00C0C000&
      Caption         =   "Contraseña de acceso"
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
      Left            =   4200
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtContra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdConAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdConCance 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label lblContra 
         BackColor       =   &H00C0C000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Fraborra 
      BackColor       =   &H00C0C000&
      Caption         =   "Contraseña de acceso"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3840
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         Height          =   495
         Left            =   1800
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtborra 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C000&
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Borrar"
      Height          =   450
      Left            =   7320
      Picture         =   "frmfactdet.frx":06F6
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Borra la  factura"
      Top             =   0
      Width           =   975
   End
   Begin MSAdodcLib.Adodc AdoFacDet 
      Height          =   330
      Left            =   3960
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoFacDet"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCamFac 
      Caption         =   "&Camb.# Fact."
      Height          =   450
      Left            =   5040
      Picture         =   "frmfactdet.frx":07E0
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Cambia el número de factura "
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   7
      Left            =   7200
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   6
      Left            =   6840
      TabIndex        =   20
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   5
      Left            =   5520
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   4
      Left            =   8160
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/mm/yy hh:mm AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7200
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtcampos 
      Height          =   615
      Index           =   0
      Left            =   1080
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin MSDataGridLib.DataGrid DbgDetVta 
      Bindings        =   "frmfactdet.frx":0922
      Height          =   4935
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
      HeadLines       =   1.5
      RowHeight       =   17
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DESGLOSE DE LA FACTURA"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "cl_producto"
         Caption         =   "clave"
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
         DataField       =   "cantidad"
         Caption         =   "Cajas"
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
      BeginProperty Column02 
         DataField       =   "cantidadp"
         Caption         =   "Piezas"
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
      BeginProperty Column03 
         DataField       =   "descripc"
         Caption         =   "                                        Descripcion"
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
      BeginProperty Column04 
         DataField       =   "medida"
         Caption         =   "Presentacion"
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
      BeginProperty Column05 
         DataField       =   "precio"
         Caption         =   "Pre. Unit."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "importe"
         Caption         =   "      Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         RecordSelectors =   0   'False
         BeginProperty Column00 
            DividerStyle    =   3
            Object.Visible         =   0   'False
            ColumnWidth     =   1514.835
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            DividerStyle    =   3
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            DividerStyle    =   3
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   3
            ColumnWidth     =   4440.189
         EndProperty
         BeginProperty Column04 
            DividerStyle    =   3
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            DividerStyle    =   3
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            DividerStyle    =   3
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   7200
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   "                                                                 Para salir presione la tecla [Esc]"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "F2  - Agegar Prod."
            TextSave        =   "F2  - Agegar Prod."
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "F3 - Quitar Prod."
            TextSave        =   "F3 - Quitar Prod."
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   450
      Left            =   8400
      Picture         =   "frmfactdet.frx":093A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Regresa a la pantalla de facturas"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   450
      Left            =   6240
      Picture         =   "frmfactdet.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancela la factura"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtcampos 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   6240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Observa  ciones"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "IEPS"
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "IVA"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbletiquetas 
      Alignment       =   2  'Center
      Caption         =   "TOTAL"
      Height          =   255
      Index           =   3
      Left            =   8160
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Factura"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Fecha cancelacion"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   13
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Motivo Cancelación"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lbletiquetas 
      Caption         =   "Serie"
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmFactDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btndes_Click()
SERIE = Mid(txtcampos(2).Text, 1, 2)
If Mid(SERIE, 2, 1) = "-" Then
    SERIE = Mid(SERIE, 1, 1)
End If
If Len(SERIE) = 1 Then
        Factura = Mid(txtcampos(2).Text, 3, Len(txtcampos(2).Text))
Else
        Factura = Mid(txtcampos(2).Text, 4, Len(txtcampos(2).Text))
End If
RESP = MsgBox("Este proceso solo Invierte la Cancelacion de la Factura, mas no la Venta de Base" & vbCrLf & "Deseas Continuar ?", vbYesNo)
If RESP = vbNo Then
   Exit Sub
End If
cn.Execute "UPDATE ventas_det SET cancelado = 0 WHERE RTRIM(SERIE) + '-'+ factura = '" & txtcampos(2).Text & "'"
rfctemp = InputBox("Escriba el RFC , para la factura..", "RFC")
'MsgBox "UPDATE facventa SET RFC = '" & Trim(rfctemp) & "',  cancelado = 0 , facfechacan = '" & Date & "', mtvocancela = ' ', cancelo = '' WHERE RTRIM(SERIE) + '-'+ numfactura = '" & Trim(txtcampos(2).Text) & "'"
cn.Execute "UPDATE facventa SET RFC = '" & Trim(rfctemp) & "',  cancelado = 0 , facfechacan = NULL, mtvocancela = ' ', cancelo = '' WHERE RTRIM(SERIE) + '-'+ numfactura = '" & Trim(txtcampos(2).Text) & "'"
cn.Execute "UPDATE facventa_det SET RFC_DET = '" & Trim(rfctemp) & "', importe = (cantidad * precio) + ( cantidadp * preciop) WHERE RTRIM(SERIE) + '-'+ ltrim(str(factura)) = '" & Trim(txtcampos(2).Text) & "'"
MsgBox "Es necesario Cambiar los importes Manualmente", vbInformation
'Me.DbgDetVta.AllowUpdate = True
End Sub

Private Sub cmbprod_LostFocus()
On Error GoTo Error:
pr = InStr(cmbprod.Text, "[")
pr1 = InStr(cmbprod.Text, "]")
strcveprod = Trim(Mid(cmbprod.Text, pr + 1, 10))
pr = InStr(cmbprod.Text, "*")
tasas(1) = Mid(cmbprod.Text, pr + 1, 1)
Select Case tasas(1)
    Case 1
        ivas(1) = 0
        iepss(1) = 0
    Case 2
        ivas(1) = 15
        iepss(1) = 0
    Case 3
        ivas(1) = 15
        iepss(1) = 25
    Case 4
        ivas(1) = 15
        iepss(1) = 30
End Select
Exit Sub
Error:
  Exit Sub
End Sub

Private Sub cmdCamFac_Click()
Dim rs As ADODB.Recordset
Dim RESP As String
fraCon.Visible = True
txtContra.SetFocus
nOp = 0
Exit Sub
Error:
    MsgBox Err.Description
    cn.RollbackTrans
End Sub

Private Sub cmdCancelar_Click()
Dim RSTFOL As New ADODB.Recordset
If Me.AdoFacturas.Recordset!cancelado Then
   MsgBox "ESTA FACTURA YA FUE CANCELADA", vbInformation
   Exit Sub
End If
If Trim(txtcampos(0).Text) = "" Then
   MsgBox "ES NECESARIO ESPECIFICAR UN MOTIVO" & Chr(13) & "PARA LA CANCELACION DE LA FACTURA", vbInformation
   txtcampos(0).SetFocus
   Exit Sub
End If
nOp = 1
fraCon.Visible = True
txtContra.SetFocus
End Sub

Private Sub cmdConAceptar_Click()
Dim RSTFOL As ADODB.Recordset
Dim RsCon As ADODB.Recordset
Dim lTrans As Boolean
On Error GoTo Error:
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT * FROM USUARIOS WHERE PASSMASTER = '" & Trim(txtContra.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
If RsCon.RecordCount = 0 Then
   MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
   txtContra.SetFocus
   SendKeys "+{HOME}"
   Exit Sub
Else
  If Not autoriza(RsCon!permisos, 2) Then Exit Sub
  fraCon.Visible = False
  txtcampos(1).Text = date & " " & Time
  'cn.BeginTrans: lTrans = True
  If cmdCancelar.Visible = True And nOp = 1 Then
     If MsgBox("DESEAS GENERAR UNA VENTA CON LA FACTURA CANCELADA", vbYesNo + vbQuestion) = vbYes Then
        Set RSTFOL = New ADODB.Recordset
        RSTFOL.Open "SELECT * FROM FOLIOS WHERE Sucursal = '" & Mid(cSucursal, 1, 3) & "' AND CAJA = '" & Caja & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
        If RSTFOL.BOF And RSTFOL.EOF Then
           'Agrego a la tabla de folios porque es primera venta de la caja
            MsgBox "BIENVENIDO AL MODULO DE VENTA A MAYOREO   " & Caja & Chr(13) & "A CONTINUACION SE REGISTRARA SU PRIMERA VENTA", vbInformation, "Bienvenida"
            cn.Execute "INSERT INTO FOLIOS(Sucursal,FolioVenta,FolioInfinito,FechaActualiza,Caja) VALUES ('" & Mid(cSucursal, 1, 3) & "',0,0,'" & date + Time & "','" & AdoFacturas.Recordset!CL_TERMINAL & "')"
        End If
        cn.Execute "UPDATE FOLIOS SET FolioVenta = FolioVenta + 1, FolioInfinito = FolioInfinito + 1, FechaActualiza = getdate() WHERE Sucursal = '" & Mid(cSucursal, 1, 3) & "' AND CAJA = '" & Trim(Mid(cCveDesUsu, 1, 3)) & "'"
        RSTFOL.Close
        RSTFOL.Open "SELECT * FROM FOLIOS WHERE Sucursal = '" & Mid(cSucursal, 1, 3) & "' AND CAJA = '" & Caja & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
        'Genero la venta
        If AdoFacturas.Recordset!globconFin Then
           cn.Execute "INSERT INTO ventas(fecha,cl_terminal,clcliente,tienda,tipoventa,folioinfinito,folioventa,chofer,agente,situacion,plaZodias,credito,prevta,modocredito,fechapagcre,montototal,Montopagos,FacRfc,folpreventa,cortey1) " & _
                       "SELECT fecha,cl_terminal,cliente = 2,tienda,tipoventa,folInf = " & RSTFOL!FolioInfinito & ",folioventa= " & RSTFOL!folioventa & " ,chofer,agente, situa = 1 ,plaZodias,credito,prevta,modocredito,fechapagcre,montot = " & AdoFacturas.Recordset!total & ",Montopagos,FacRfc,folpreventa,CorteY1 FROM VENTAS WHERE noventa = " & AdoFacturas.Recordset!noventa
        Else
           cn.Execute "INSERT INTO ventas(fecha,cl_terminal,clcliente,tienda,tipoventa,folioinfinito,folioventa,chofer,agente,situacion,plaZodias,credito,prevta,modocredito,fechapagcre,montototal,Montopagos,FacRfc,folpreventa,cortey1) " & _
                       "SELECT fecha,cl_terminal,clcliente,tienda,tipoventa,folInf = " & RSTFOL!FolioInfinito & ",folioventa= " & RSTFOL!folioventa & " ,chofer,agente, situa = 1 ,plaZodias,credito,prevta,modocredito,fechapagcre,montot = " & AdoFacturas.Recordset!total & " ,Montopagos,FacRfc,folpreventa,CorteY1 FROM VENTAS WHERE noventa = " & AdoFacturas.Recordset!noventa
        End If
        RSTFOL.Close
        RSTFOL.Open "SELECT MAX(NOVENTA) AS FolVta FROM VENTAS", cn, adOpenKeyset, adLockOptimistic, adCmdText
        'Genero nuevamente una venta con el detalle de la factura
        If AdoFacturas.Recordset!globconFin Then
           'cn.Execute "INSERT INTO VENTAS_DET(noventa,cl_producto,cantidad,cantidadp,precio,precosto,importe,ieps,iva) SELECT  FolVta = " & RSTFOL!folVta & ", MAX(cl_producto), SUM(cantidad),SUM(cantidadp), AVG(precio), AVG(precosto), SUM(importe), MAX(ieps), MAX(iva) FROM ventas_det WHERE RTRIM(SERIE) + '-'+ factura = '" & txtcampos(2).Text & "' GROUP BY CL_PRODUCTO"
           'MsgBox "INSERT INTO VENTAS_DET(noventa,cl_producto,cantidad,cantidadp,precio,precosto,importe,ieps,iva,tipocantidad,cancelado,tasaieps,precostop,preciop) SELECT  FolVta = " & RSTFOL!folVta & ", producto, cantidad, cantidadp, precio, costo, importe, ieps, iva, tipo = 1, cancel  = 0, tasaieps, preciop, costop) FROM facventa_det WHERE RTRIM(SERIE) + '-'+ factura = '" & txtcampos(2).Text & "'"
           cn.Execute "INSERT INTO VENTAS_DET(noventa,cl_producto,cantidad,cantidadp,precio,precosto,importe,ieps,iva,tipocantidad,cancelado,tasaieps,precostop,preciop) SELECT  FolVta = " & RSTFOL!folVta & ", producto, cantidad, cantidadp, precio, costo, importe, ieps, iva, tipo = 1, cancel  = 0, tasaieps, preciop, costop FROM facventa_det WHERE RTRIM(SERIE) + '-'+ factura = '" & txtcampos(2).Text & "'"
        Else
           cn.Execute "INSERT INTO VENTAS_DET(noventa,cl_producto,cantidad,cantidadp,precio,precosto,importe,ieps,iva,tipocantidad,cancelado,tasaieps,precostop,preciop) SELECT FolVta = " & RSTFOL!folVta & ", cl_producto,cantidad,cantidadp,precio,precosto,importe,ieps,iva,tipocantidad, cancel = 0,tasaieps,precostop,preciop FROM ventas_det WHERE RTRIM(SERIE) + '-'+ factura = '" & txtcampos(2).Text & "'"
        End If
        MsgBox "SE GENERO NUEVAMENTE LA VENTA " & Chr(13) & Chr(13) & "CON FOLIO UNICO " & RSTFOL!folVta, vbInformation, "Area de facturación"
    Else
        cn.Execute "UPDATE inventario SET incant = incant + cantidad, InCantPza = InCantPza + cantidadp FROM Ventas_det WHERE Inprod = cl_producto AND RTRIM(SERIE) + '-'+ factura = '" & txtcampos(2).Text & "'"
    End If
    cn.Execute "UPDATE ventas_det SET cancelado = 1, importe = 0 WHERE RTRIM(SERIE) + '-'+ factura = '" & txtcampos(2).Text & "'"
    cn.Execute "UPDATE facventa SET RFC = 'CANC999999999', Total = 0, cancelado = 1 , facfechacan = '" & txtcampos(1).Text & "', mtvocancela = '" & IIf(IsNull(txtcampos(0).Text), ".", txtcampos(0).Text) & "', cancelo = '" & RsCon!login & " ' WHERE RTRIM(SERIE) + '-'+ numfactura = '" & txtcampos(2).Text & "'"
    'PARA QUE SALGA EN CEROS
    'SERIE = Mid(txtcampos(2).Text, 1, 2)
    'If Mid(SERIE, 2, 1) = "-" Then
    '    SERIE = Mid(SERIE, 1, 1)
    'End If
    'If Len(SERIE) = 1 Then
    '    Factura = Mid(txtcampos(2).Text, 3, Len(txtcampos(2).Text))
    'Else
    '    Factura = Mid(txtcampos(2).Text, 4, Len(txtcampos(2).Text))
    'End If
    nSepara = InStr(1, txtcampos(2).Text, "-")
    SERIE = Mid(txtcampos(2).Text, 1, nSepara - 1)
    Factura = Trim(Mid(txtcampos(2).Text, nSepara + 1, Len(txtcampos(2).Text)))
    
    cn.Execute "UPDATE facventa_det SET rfc_det = 'CANC999999999', importe = 0 WHERE factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
    cn.Execute "UPDATE ventas SET montototal = (SELECT SUM(importe) FROM ventas_det WHERE noventa = " & AdoFacturas.Recordset!noventa & ") WHERE noventa = " & AdoFacturas.Recordset!noventa
    cn.Execute "DELETE FROM abonos WHERE SERIE = '" & Trim(SERIE) & "' AND FACTURA = '" & Trim(Factura) & "'"
  ElseIf nOp = 0 Then
        RESP = InputBox("Introduzca el nuevo numero de factura", "Cambiar el numero de factura")
        If Trim(RESP) = "" Then
           Exit Sub
        End If
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM FACVENTA WHERE serie = '" & Trim(AdoFacturas.Recordset!SERIE) & "' AND numfactura = '" & Trim(RESP) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
        If rs.BOF And rs.EOF Then
            cn.Execute "UPDATE VENTAS_DET SET factura = '" & Trim(RESP) & "' WHERE SERIE = '" & AdoFacturas.Recordset!SERIE & "' AND FACTURA = '" & AdoFacturas.Recordset!numfactura & "'"
            cn.Execute "UPDATE FACVENTA SET numfactura = '" & Trim(RESP) & "' WHERE facclave = " & AdoFacturas.Recordset!facclave
            cn.Execute "UPDATE FACVENTA_DET SET factura = " & Val(Trim(RESP)) & " WHERE SERIE ='" & Trim(AdoFacturas.Recordset!SERIE) & "' AND FACTURA = '" & Trim(AdoFacturas.Recordset!numfactura) & "'"
            cn.Execute "UPDATE abonos SET factura = '" & Trim(RESP) & "' WHERE SERIE ='" & Trim(AdoFacturas.Recordset!SERIE) & "' AND FACTURA = '" & Trim(AdoFacturas.Recordset!numfactura) & "'"
        Else
            MsgBox "YA EXISTE EL NUMERO DE FACTURA EN LA SERIE " & AdoFacturas.Recordset!SERIE & Chr(13) & "ASOCIADA A LA VENTA CON FOLIO UNICO " & rs!noventa, vbExclamation
            Exit Sub
        End If
  End If
  Unload Me
End If
Exit Sub
Error:
MsgBox Err.Description
'If lTrans Then cn.RollbackTrans
End Sub

Private Sub cmdConCance_Click()
   fraCon.Visible = False
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo Error:
  AdoFacturas.Recordset!total = Format(txtcampos(4).Text, "#############.##")
  AdoFacturas.Recordset!TotFac = Format(txtcampos(4).Text, "#############.##")
  AdoFacturas.Recordset.Update
  Unload Me
  Exit Sub
Error:
  MsgBox Err.Description
End Sub

Private Sub cmdRegresar_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
'SE BORRA LA FACTURA
RESP = MsgBox("Realmente deseas Borrar la Factura ? ", vbYesNo + vbQuestion)
If RESP = vbYes Then
    Fraborra.Enabled = True
    Fraborra.Visible = True
    txtborra.SetFocus
End If
End Sub

Private Sub Command3_Click()
prod = Me.AdoFacDet.Recordset!producto
SERIE = Mid(txtcampos(2).Text, 1, 2)
If Mid(SERIE, 2, 1) = "-" Then
   SERIE = Mid(SERIE, 1, 1)
End If
If Len(SERIE) = 1 Then
    Factura = Mid(txtcampos(2).Text, 3, Len(txtcampos(2).Text))
Else
    Factura = Mid(txtcampos(2).Text, 4, Len(txtcampos(2).Text))
End If
CAD = "DELETE FACVENTA_det WHERE FACTURA = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' and producto = '" & Trim(prod) & "'"
cn.Execute CAD
fracambia.Enabled = False
fracambia.Visible = False
End Sub

Private Sub Command4_Click()
Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT * FROM USUARIOS WHERE PASSMASTER = '" & Trim(txtborra.Text) & "'", cn, adOpenKeyset, adLockOptimistic, adCmdText
If RsCon.RecordCount = 0 Then
   MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
   txtborra.SetFocus
   SendKeys "+{HOME}"
   Exit Sub
Else
   If Not autoriza(RsCon!permisos, 2) Then Exit Sub
   'SERIE = Mid(txtcampos(2).Text, 1, 2)
   'If Mid(SERIE, 2, 1) = "-" Then
   'SERIE = Mid(SERIE, 1, 1)
   'End If
   'If Len(SERIE) = 1 Then
   '   Factura = Mid(txtcampos(2).Text, 3, Len(txtcampos(2).Text))
   'Else
   '   Factura = Mid(txtcampos(2).Text, 4, Len(txtcampos(2).Text))
   'End If
   nSepara = InStr(1, txtcampos(2).Text, "-")
   SERIE = Mid(txtcampos(2).Text, 1, nSepara - 1)
   Factura = Trim(Mid(txtcampos(2).Text, nSepara + 1, Len(txtcampos(2).Text)))
   
   cn.Execute "Delete facventa_det where factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
   cn.Execute "Delete facventa where numfactura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "'"
   cn.Execute "UPDATE ventas_det SET factura = null, serie = null, facturado = 0 WHERE serie = '" & Trim(SERIE) & "' and factura = '" & Trim(Factura) & "'"
   cn.Execute "UPDATE VENTAS set situacion = 1 where noventa = " & Me.AdoFacturas.Recordset!noventa
   MsgBox "Se borro la Factura...", vbInformation
   Fraborra.Enabled = False
   Fraborra.Visible = False
   RsCon.Close
   Set RsCon = Nothing
End If
End Sub

Private Sub Command5_Click()
'se carga el catalogo de productos
cmbprod.Clear
Dim RsCon As ADODB.Recordset
Set RsCon = New ADODB.Recordset
RsCon.Open "SELECT consec,descripc,contenid,medida,paquetes,tasaieps,fecact FROM tfproduc order by descripc ", cn, adOpenDynamic, adLockOptimistic, adCmdText
If RsCon.EOF Then
   MsgBox "No existen Productos en el catalogo"
Else
   While Not RsCon.EOF
   If Not (IsNull(RsCon!medida)) And Not IsNull(RsCon!PAQUETES) Then
                 cmbprod.AddItem RsCon!descripc + " ( " + RsCon!medida + " ) " _
                  + " [ " + RsCon!CONSEC + " ]*" & RsCon!tasaieps
   End If
   RsCon.MoveNext
   Wend
End If
End Sub

Private Sub DbgDetVta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmnewprod.ventas.Caption = txtcampos(7).Text
   frmnewprod.Factura.Caption = txtcampos(2).Text
   If Not (AdoFacDet.Recordset.BOF And AdoFacDet.Recordset.EOF) Then
        frmnewprod.Caption = AdoFacDet.Recordset!descripc + " " + AdoFacDet.Recordset!medida
        frmnewprod.txtcajas.Text = Me.AdoFacDet.Recordset!cantidad
        frmnewprod.txtpzas.Text = AdoFacDet.Recordset!cantidadp
        frmnewprod.txtuni.Text = AdoFacDet.Recordset!PRECIO
        frmnewprod.txttot.Text = AdoFacDet.Recordset!importe
    Else
    End If
    frmnewprod.fracambia.Refresh
    frmnewprod.Show 1
End If
End Sub

Private Sub Form_Activate()
SendKeys vbTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case 13
      'KeyAscii = 0
      'SendKeys vbTab
 Case 27
      Unload Me
 Case 113 ' F2
      'para agregar un producto
      
End Select
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub txtcajas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtpzas.SetFocus
End If
End Sub

Private Sub txtborra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub

Private Sub txtCampos_LostFocus(Index As Integer)
Dim rstt As ADODB.Recordset
Select Case Index
Case 2
    'Cargo los productos de la factura
    AdoFacDet.ConnectionString = cn
    nSepara = InStr(1, txtcampos(2).Text, "-")
    SERIE = Mid(txtcampos(2).Text, 1, nSepara - 1)
    Factura = Trim(Mid(txtcampos(2).Text, nSepara + 1, Len(txtcampos(2).Text)))
    CAD = " select facventa_det.ieps as ieps, facventa_det.iva as iva, facventa_det.tasaieps as tasaieps, cantidad,cantidadp,precio, importe, LTRIM(STR(paquetes)) + ' X ' + LTRIM(STR(contenid,10,3)) + ' ' + medida AS MEDIDA, descripc as DESCRIPC, producto from facventa_det,tfproduc where consec = producto and factura = '" & Trim(Factura) & "' and serie = '" & Trim(SERIE) & "' ORDER BY descripc"
    
    AdoFacDet.RecordSource = CAD
    AdoFacDet.Refresh
    AdoFacturas.ConnectionString = cCadConex
    CAD = "SELECT total, totfac, globconfin, noventa, serie,numfactura,facclave, facfechacan,mtvocancela,total,iva, ieps,cancelado,concepto FROM FACVENTA WHERE serie = '" & Trim(SERIE) & "' and numfactura = '" & Trim(Factura) & "'"
    AdoFacturas.RecordSource = CAD
    AdoFacturas.Refresh
    If AdoFacturas.Recordset.RecordCount > 0 Then
        txtcampos(1).Text = IIf(IsNull(AdoFacturas.Recordset!FACFECHACAN), "", AdoFacturas.Recordset!FACFECHACAN)
        cmdCancelar.Enabled = True
        txtcampos(5).Text = Format(AdoFacturas.Recordset!iva, "$ ###,###,##0.00")
        txtcampos(6).Text = Format(AdoFacturas.Recordset!ieps, "$ ###,###,##0.00")
        On Error Resume Next
        txtcampos(7).Text = AdoFacturas.Recordset!noventa
        txtcampos(0).Text = IIf(IsNull(AdoFacturas.Recordset!MtvoCancela), "", Trim(AdoFacturas.Recordset!MtvoCancela))
        Set rstt = New ADODB.Recordset
        rstt.Open "SELECT SUM(importe) as impte from facventa_det where factura = " & AdoFacturas.Recordset!numfactura & " and serie = '" & AdoFacturas.Recordset!SERIE & "'", cn, adOpenForwardOnly, adLockOptimistic, adCmdText
        txtcampos(4).Text = Format(rstt!impte, "$ ###,###,##0.00")
        rstt.Close
    End If
End Select
End Sub

Private Sub txtpzas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtuni.SetFocus
End If
End Sub

Private Sub txttot_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btncambia.SetFocus
End If
End Sub

Private Sub txtuni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txttot.SetFocus
End If
End Sub

Private Sub txtContra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys vbTab
   KeyAscii = 0
End If
End Sub
