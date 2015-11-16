VERSION 5.00
Begin VB.Form FrmReportes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   2175
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Opcion 
      Caption         =   "Reporte de Conteos"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox TxtInventario 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Clave del Inventario"
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton Opcion 
      Caption         =   "Acta de entrega"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   3120
      MouseIcon       =   "FrmReportes.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Registra el inventario"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3120
      MouseIcon       =   "FrmReportes.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Regresa al menu principal"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Inventario:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   750
   End
End
Attribute VB_Name = "FrmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     'Las declaraciones de variables seran obligatorias

Private Sub CmdAceptar_Click()
    'Acta de entrega
    If Opcion(0).Value = True Then ActaDeEntrega
    
    'Reporte de Conteos
    If Opcion(1).Value = True Then ReporteConteos

End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    'Si Hay un inventario en proceso
    If INVENTARIO <> "" Then
        TxtInventario = INVENTARIO
    End If
End Sub

Sub ReporteConteos()
Dim PARRAFO1 As String
Dim PARRAFO2 As String
Dim PARRAFO3 As String
Dim PARRAFO4 As String
Dim PARRAFO5 As String
'Dim TABLA As New ADODB.Recordset
Dim SQL As String

    If TxtInventario = "" Then Exit Sub
    
    DE1.Commands("ReporteConteo").Parameters("Inventario").Value = TxtInventario
    If DE1.rsReporteConteo.State = adStateClosed Then DE1.rsReporteConteo.Open

    If DE1.rsReporteConteo.BOF And DE1.rsReporteConteo.EOF Then
        MsgBox "NO HAY DATOS AUN", vbExclamation
        DE1.rsReporteConteo.Close
    Else
        Reporte1.Show    'Muestro el Reporte
    End If
    If DE1.rsReporteConteo.State = adStateOpen Then DE1.rsReporteConteo.Close
End Sub

Sub ActaDeEntrega()
Dim PARRAFO1 As String
Dim PARRAFO2 As String
Dim PARRAFO3 As String
Dim PARRAFO4 As String
Dim PARRAFO10 As String

Dim TABLA As New ADODB.Recordset
Dim SQL As String
Dim VALOR As Currency   'VALOR TOTAL DE LA MERCANCIA DEL INVENTARIO ACTUAL

    INVENTARIO = TxtInventario
    If INVENTARIO = "" Then Exit Sub
    
    DE1.Commands("ActaEntrega").Parameters("Inventario").Value = TxtInventario
    If DE1.rsActaEntrega.State = adStateClosed Then DE1.rsActaEntrega.Open

    If DE1.rsActaEntrega.BOF And DE1.rsActaEntrega.EOF Then
        MsgBox "NO HAY DATOS AUN", vbExclamation
        DE1.rsActaEntrega.Close
    Else
        
'        PARRAFO1 = "EN LA CIUDAD DE " & DE1.rsActaEntrega!MUNICIPIO & " CHIAPAS SE LEVANTA LA PRESENTE ACTA DE ENTREGA, SIENDO LAS " & Format(DE1.rsActaEntrega!HORA_INI, "HH:MM") & " HORAS DEL DÍA " & _
'        UCase(Format(DE1.rsActaEntrega!FECHA_INI, "d "" de "" MMMM "" DEL AÑO "" yyyy")) & ", EN LAS INSTALACIONES DE LA EMPRESA HOLDING MÉXICO CENTRO AMÉRICA S. A. DE C. V. " & _
'        " SUCURSAL " & DE1.rsActaEntrega!Desc_Tienda & ", UBICADA EN " & DE1.rsActaEntrega!DirTda ' & " DE ESTA CIUDAD DE " & DE1.rsActaEntrega!MUNICIPIO & ", CHIAPAS.  ESTANDO PRESENTE LAS SIGUIENTES PERSONAS, "
        PARRAFO1 = "EN LA CIUDAD DE " & DE1.rsActaEntrega!MUNICIPIO & " CHIAPAS SE LEVANTA LA PRESENTE ACTA DE ENTREGA," & _
        " SIENDO LAS " & Format(Time, "HH:MM") & " HORAS DEL DÍA " & UCase(Format(Date, "d "" de "" MMMM "" DEL AÑO "" yyyy")) & ", EN LAS INSTALACIONES DE LA EMPRESA HOLDING MÉXICO CENTRO AMÉRICA S. A. DE C. V. " & _
        " SUCURSAL " & DE1.rsActaEntrega!DescTda & ", UBICADA EN " & DE1.rsActaEntrega!DirTda & " DE ESTA CIUDAD DE " & DE1.rsActaEntrega!MUNICIPIO & ", CHIAPAS.  ESTANDO PRESENTE LAS SIGUIENTES PERSONAS, "
        
        SQL = "SELECT F.FUNCION, E.NOMBRE FROM INVENTARIO I, RESPONSABLES R, FUNCIONES F, EMPLEADOS E WHERE F.CLAVE = R.FUNCION AND R.CLAVE = E.CLAVE AND R.CLAVE = I.RESP_INVENTARIO AND F.CLAVE = 'RESINV' AND I.CLAVE = '" & INVENTARIO & "'"
        Set TABLA = Conn.Execute(SQL)
        PARRAFO1 = PARRAFO1 & TABLA.Fields("NOMBRE") & " COMO " & TABLA.Fields("FUNCION") & " Y EL  C. "
        
        SQL = "SELECT F.FUNCION, E.NOMBRE FROM INVENTARIO I, RESPONSABLES R, FUNCIONES F, EMPLEADOS E WHERE I.RESP_TIENDA = R.CLAVE AND R.FUNCION = F.CLAVE AND R.CLAVE = E.CLAVE AND F.CLAVE = 'RESTDA' AND I.CLAVE = '" & INVENTARIO & "'"
        Set TABLA = Conn.Execute(SQL)
        PARRAFO1 = PARRAFO1 & TABLA.Fields("NOMBRE") & " COMO " & TABLA.Fields("FUNCION") & ",  PARA HACER CONSTAR LOS SIGUIENTES HECHOS:"
        
        PARRAFO2 = "PRIMERO: CON FECHA " & UCase(Format(DE1.rsActaEntrega!FECHA_INI, "d "" de "" MMMM "" DEL AÑO "" yyyy"))
        PARRAFO2 = PARRAFO2 & ", SE PROCEDIO A LEVANTAR INVENTARIO FISICO, A LA VISTA TODOS Y CADA UNO DE LOS PRODUCTOS  QUE INTEGRAN  LA  BODEGA, EN PRESENCIA   DEL C. " & TABLA.Fields("NOMBRE") & " COMO " & TABLA.Fields("FUNCION") & "; Y AUXILIARES DE BODEGA"
        
        PARRAFO3 = "SEGUNDO: DESPUES DEL INVENTARIO GENERAL SE HA DADO SEGUIMIENTO PARA CHECAR QUE LA DOCUMENTACIÓN DE ENTRADAS Y ENVIOS SE REALICEN DE FORMA CORRECTA, ASI TAMBIEN SE ESTAN REALIZANDO INVENTARIO DE FORMA ALEATORIA, LO CUAL EL RESULTADO HA SIDO SATISFACTORIO, ESTO ES CON  LA FINALIDAD DE DAR SEGUIMIENTO QUE EL INVENTARIO ESTE FUNCIONANDO CON CERTEZA DEL 100%."
        
        PARRAFO4 = "TERCERO: CONSCIENTES DE QUE EL SISTEMA DE INVENTARIO FUNCIONA DE MANERA CORRECTA, SE LE HACE ENTREGA AL C. " & _
        TABLA.Fields("NOMBRE") & " COMO " & TABLA.Fields("FUNCION") & " LA BODEGA CON LOS PRODUCTOS CORRESPONDIENTES, POR LO TANTO A PARTIR DE ESTA FECHA SE HACEN RESPONSABLES DE SU BUEN FUNCIONAMIENTO, RESGUARDANDO Y CUSTODIANDO TODOS Y CADA  UNO DE LOS PRODUCTOS QUE AQUÍ SE CONTROLAN DEL CUAL SE FIRMA CADA UNA DE  LAS HOJAS DONDE SE REFLEJA EL INVENTARIO POR EXISTENCIAS Y VALORIZADO CADA UNO DE LOS PRODUCTOS, QUE EN SU MOMENTO TUVO A LA VISTA CUANTIFICADO DE LA SIGUIENTE MANERA."
        'Si solo se ha terminado el primer conteo se toma ese valor, si no se toma el valor del segundo conteo
        VALOR = IIf(DE1.rsActaEntrega!CONTEO1 = 0, DE1.rsActaEntrega!CONTEO2, DE1.rsActaEntrega!CONTEO1)
        
        PARRAFO4 = PARRAFO4 & " VALOR  TOTAL DE LA MERCANCIA " & Format(VALOR, "$  ###,###,###.00") & " (" & NumLet$(Format(VALOR, "#########0.00")) & ")"

        PARRAFO10 = "PARA LOS EFECTOS LEGALES A QUE TUVIERA LUGAR Y CONCIENTES DE SU CONTENIDO, SE CIERRA LA PRESENTE FIRMANDO AL CALCE Y MARGEN CADA UNA DE SUS HOJAS PARA DAR FE, SIENDO LAS " & _
        Format(DE1.rsActaEntrega!HORA_FIN, "HH:MM") & " HORAS  DEL  DIA " & UCase(Format(DE1.rsActaEntrega!FECHA_FIN, "d "" de "" MMMM "" DEL AÑO "" yyyy")) & "."
                
        With ActaEntrega2
            .Sections("Texto").Controls.Item("LblParrafo1").Caption = PARRAFO1
            .Sections("Texto").Controls.Item("LblParrafo2").Caption = PARRAFO2
            .Sections("Texto").Controls.Item("LblParrafo3").Caption = PARRAFO3
            .Sections("Texto").Controls.Item("LblParrafo4").Caption = PARRAFO4
            .Sections("Texto").Controls.Item("LblParrafo10").Caption = PARRAFO10
        End With
        
        ActaEntrega2.Show    'Muestro el Reporte
            
    End If
    
    If DE1.rsActaEntrega.State = adStateOpen Then DE1.rsActaEntrega.Close
    
End Sub

'Private Sub TxtInventario_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'End Sub
