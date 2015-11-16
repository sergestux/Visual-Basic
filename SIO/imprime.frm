VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "IMPRIME.frx":0000
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3625
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ventas"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ventas"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblventa_d"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()

Set rs = New ADODB.Recordset
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    rs.Source = "select clafamil, sum(precio_v) as suma from tfproduc,tblventa_d where consec = pclave group by clafamil"
    rs.ActiveConnection = cn
    rs.Open
    Set Adodc1.Recordset = rs
    Adodc1.Refresh
    rs.Close
    rs.Source = "select min(ticket) as minimo,max(ticket) as maximo from tblventa_d"
    rs.Open
    
Dim total As Double
Dim dondex  As Integer
Call imprimet
Const donde = 150

Printer.ScaleMode = vbPoints
Printer.Print " "
Printer.Print " "
Printer.Print "      Corte de Caja X"
Printer.Print "     " + Format(Date, "dd/mm/yy") + " " + Str(Time)
Printer.Print " "
Printer.Print "Del ticket" & Str(rs.Fields!minimo) & " al " & Str(rs.Fields!maximo)
Printer.Print " "
Adodc1.Recordset.MoveFirst
total = 0
For i = 1 To Adodc1.Recordset.RecordCount
Printer.Print "Depto " + Adodc1.Recordset!clafamil;
dondex = donde - (50 + (Printer.TextWidth(Format(Adodc1.Recordset!suma, "###,###,##0.00"))))
dondex = 50 + dondex
Printer.CurrentX = dondex
Printer.Print Format(Adodc1.Recordset!suma, "###,###,##0.00")
total = total + Adodc1.Recordset!suma
Adodc1.Recordset.MoveNext
Next
Printer.Print " "
Printer.Print " Total $";

dondex = donde - (50 + (Printer.TextWidth(Format(total, "###,###,##0.00"))))
dondex = 50 + dondex

Printer.CurrentX = dondex
Printer.Print Format(total, "###,###,##0.00")
Printer.Print "------------------------------------------"
Printer.Print "  Efectivo  "
Printer.Print "    Vales  "
Printer.Print "   Cheque  "
Printer.Print "T.Credito  "
Printer.EndDoc

End Sub

Private Sub Command2_Click()
    Call ETIQUETAS
    
   End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
    cn.ConnectionString = "DSN=ventas"
    cn.Open
Printer.ScaleMode = vbPoints
    
End Sub

Private Sub imprimet()
Dim nprodtot As Integer
Const donde = 150
    
With Printer.Font
       .Name = "Arial Narrow"
       .Size = 8
       .Bold = False
End With


Printer.ScaleMode = vbPoints
Printer.Print " "
Printer.Print " "

dondex = Printer.TextWidth("VINOS Y LICORES") / 2
Printer.CurrentX = 75 - dondex
Printer.Print "VINOS Y LICORES"

dondex = Printer.TextWidth("SUCURSAL REFORMA") / 2
Printer.CurrentX = 75 - dondex
Printer.Print "SUCURSAL REFORMA"



dondex = Printer.TextWidth("Cajera:") / 2
Printer.CurrentX = 75 - dondex
Printer.Print "Cajera: "

'dondex = Printer.TextWidth("Ticket " & Str(rs.Fields!minimo)) / 2
'Printer.CurrentX = 75 - dondex
'Printer.Print "Ticket " & Str(rs.Fields!minimo)

Printer.Print " "
Adodc1.Recordset.MoveFirst
total = 0
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
nprodtot = Adodc1.Recordset.RecordCount
For i = 0 To nprodtot - 1

'imprime clave del producto y descripcion
Printer.Print DataGrid1.Columns(1).Text, DataGrid1.Columns(2).Text;
'imprime el precio
dondex = donde - (50 + (Printer.TextWidth(Format(DataGrid1.Columns(3).Text, "###,###,##0.00"))))
dondex = 50 + dondex
Printer.CurrentX = dondex
Printer.Print Format(DataGrid1.Columns(3).Text, "###,###,##0.00")
total = total + Val(DataGrid1.Columns(3).Text)

Adodc1.Recordset.MoveNext

Next
Printer.Print " "
Printer.Print " Importe $";
dondex = donde - (50 + (Printer.TextWidth(Format(total, "###,###,##0.00"))))
dondex = 50 + dondex
Printer.CurrentX = dondex
Printer.Print Format(total, "###,###,##0.00")
Printer.Print " "
Printer.Print "Total de Productos vendidos  " + Str(nprodtot)
Printer.Print " "
dondex = Printer.TextWidth(Format(Date, "dd/mm/yy") + " " + Str(Time)) / 2
Printer.CurrentX = 75 - dondex
Printer.Print Format(Date, "dd/mm/yy") + " " + Str(Time)
dondex = Printer.TextWidth("*** GRACIAS POR SU COMPRA ***") / 2
Printer.Print " "
Printer.CurrentX = 75 - dondex
Printer.Print "*** GRACIAS POR SU COMPRA ***"

For i = 1 To 3
Printer.Print " "
Next

Printer.EndDoc
End Sub


Private Sub ETIQUETAS()
Set rs = New ADODB.Recordset
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    rs.Source = "SELECT consec,descripc,costocaj  FROM tfproduc"
    rs.ActiveConnection = cn
    rs.Open
rs.MoveFirst

For i = 1 To 50

For j = 1 To 2
Printer.Print " "
Next

With Printer.Font
        .Name = "Arial Narrow"
        .Size = 8
        .Bold = False
    End With
    
    Printer.Print rs.Fields!consec;
    rs.MoveNext
    Printer.CurrentX = 200
    Printer.Print rs.Fields!consec;
    rs.MoveNext
    Printer.CurrentX = 400
    Printer.Print rs.Fields!consec
For j = 1 To 2
rs.MovePrevious
Next
    Printer.Print rs.Fields!descripc;
    rs.MoveNext
    Printer.CurrentX = 200
    Printer.Print rs.Fields!descripc;
    rs.MoveNext
    Printer.CurrentX = 400
    Printer.Print rs.Fields!descripc


For j = 1 To 2
rs.MovePrevious
Next
    
    With Printer.Font
        .Name = "Arial"
        .Size = 14
        .Bold = True
    End With
    
    Printer.Print Format(rs.Fields!costocaj, "###,###,##0.00");
    rs.MoveNext
    Printer.CurrentX = 200
    Printer.Print Format(rs.Fields!costocaj, "###,###,##0.00");
    rs.MoveNext
    Printer.CurrentX = 400
    Printer.Print Format(rs.Fields!costocaj, "###,###,##0.00")
    
    rs.MoveNext
    
Next
Printer.EndDoc
End Sub
