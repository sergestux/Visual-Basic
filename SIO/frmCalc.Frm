VERSION 5.00
Begin VB.Form frmCalc 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Calculadora            [PITICO]"
   ClientHeight    =   5265
   ClientLeft      =   3600
   ClientTop       =   1560
   ClientWidth     =   4185
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "g32Calc.frx":0000
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5265
   ScaleWidth      =   4185
   Begin VB.Timer tim1 
      Interval        =   100
      Left            =   2760
      Top             =   2880
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   4470
      Width           =   705
   End
   Begin VB.ListBox LstMem 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   60
      TabIndex        =   34
      Top             =   390
      Width           =   1755
   End
   Begin VB.CommandButton CmdOff 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1140
      TabIndex        =   33
      Top             =   3150
      Width           =   1005
   End
   Begin VB.VScrollBar scrMem 
      Height          =   465
      Left            =   570
      Max             =   9
      TabIndex        =   32
      Top             =   2580
      Width           =   255
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   3570
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4860
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   2220
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4860
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4860
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4470
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   3570
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4470
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3690
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "1/X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   3570
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3690
      Width           =   495
   End
   Begin VB.CommandButton cmdDecimal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4860
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1140
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4860
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1140
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4470
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   1680
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4470
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   2220
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4470
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   2220
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   1680
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   1140
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   1140
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3690
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3690
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   2220
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3690
      Width           =   495
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   4860
      Width           =   705
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   705
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   3690
      Width           =   705
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   3300
      Width           =   705
   End
   Begin VB.CommandButton cmdPercent 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3570
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3570
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3150
      Width           =   495
   End
   Begin VB.CommandButton cmdCancelEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3150
      Width           =   495
   End
   Begin VB.CommandButton CmdBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2220
      TabIndex        =   27
      Top             =   3150
      Width           =   495
   End
   Begin VB.ListBox LstTira 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2160
      Left            =   1890
      TabIndex        =   29
      Top             =   390
      Width           =   2235
   End
   Begin VB.PictureBox pctKeyTrap 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   60
      ScaleHeight     =   270
      ScaleWidth      =   2730
      TabIndex        =   0
      Top             =   5355
      Width           =   2760
   End
   Begin VB.Label lbleti 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   840
      TabIndex        =   37
      Top             =   90
      Width           =   3255
   End
   Begin VB.Label LblMenu 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Archivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   36
      Top             =   75
      Width           =   735
   End
   Begin VB.Label lblReadout 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   960
      TabIndex        =   1
      Top             =   2670
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Mem:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2580
      Width           =   435
   End
   Begin VB.Label LblMem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   28
      Top             =   2820
      Width           =   435
   End
   Begin VB.Menu mnuArc 
      Caption         =   "&Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnuArcLee 
         Caption         =   "&Leer"
      End
      Begin VB.Menu mnuArcGua 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnuArcSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdiTir 
         Caption         =   "Borrar &Tira"
      End
      Begin VB.Menu mnuBorMem 
         Caption         =   "Borrar &Memorias"
      End
      Begin VB.Menu mnuArcSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimizar 
         Caption         =   "Mi&nimizar"
      End
      Begin VB.Menu mnuArcSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "&Acerca de..."
      End
      Begin VB.Menu mnuArcSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcSalir 
         Caption         =   "&Salir"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
' Visual Basic for MS-DOS Calculator Program
' Calculator program that supports mouse and keyboard
' input.
' Copyright (C) 1982-1992 Microsoft Corporation
' You have a royalty-free right to use, modify, reproduce
' and distribute the sample applications and toolkits provided with
' Visual Basic for MS-DOS (and/or any modified version)
' in any way you find useful, provided that you agree that
' Microsoft has no warranty, obligations or liability for
' any of the sample applications or toolkits.
' ------------------------------------------------------------------------
' Versi¢n 2.0  para Visual Basic DOS    ( 2/Abr/96)
' Ampliado el funcionamiento básico que tenía, con
' 10 Memorias y Lista de los últimos números...
' Revisión para Windows95               ( 9/Abr/96)
' Última revisión:                      ( 9/Jul/96)
' -------------------------------------------------
'DefInt A-Z
Option Explicit

Const ANCHOMEM = 14
Const ANCHOTIRA = 16

Dim Op1 As Double           ' Previously input operand.
Dim Op2 As Double           ' Second operand.
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Number of operands.
Dim LastInput As String     ' Indicate type of last keypress.
Dim OpFlag As String        ' Indicate pending operation.
Dim Decimal1 As String       ' Decimal indicator (country specific).
Dim nList As Integer        ' N£mero de items en LstTira
Dim EnList As Integer       ' Modificando el List
Dim MaxItemEnList As Integer
Dim frmT, frmL              ' Posiciones del Form
Dim frmH, frmW
Dim nPos  As Integer
Dim nPos1  As Integer

Private Sub CmdBack_Click()
    Dim temp$
    Dim c$
    Dim i%

    temp$ = LTrim$(RTrim$(lblReadout.Caption))
    i% = Len(temp$)
    If i > 2 Then
        If Right$(temp$, 1) = Decimal1 Then
            temp$ = Left$(temp$, i - 2) + Decimal1
        Else
            'Quitar el último caracter
            temp$ = Left$(temp$, i - 1)
        End If
        If Decimal1 = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
        lblReadout.Caption = Format(Val(temp$), "#.####")
    Else
        DecimalFlag = False
        lblReadout.Caption = Format(0, "0.")
        LastInput = "OPS"
        Beep
    End If
End Sub

' Click event procedure for Cancel (C) key.
' Reset the display and initializes variables.
'
Private Sub cmdCancel_Click()
    lblReadout.Caption = Format(0, "0.")
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = ""
    pctKeyTrap.SetFocus
End Sub

' Click event procedure for Cancel Entry (CE) key.
'
Private Sub cmdCancelEntry_Click()
    lblReadout.Caption = Format(0, "0.")
    DecimalFlag = False                 ' Decimal point here does not indicate fraction.
    LastInput = "CE"                    ' Track what user did last.
    pctKeyTrap.SetFocus
End Sub

' Click event procedure for Decimal key.
' If last keypress was an operator, initialize
' readout. Otherwise, append a decimal
' point to the display.
'
Private Sub cmdDecimal_Click()
    If LastInput = "NEG" Then
        lblReadout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        'lblReadOut.Caption = "0" + Decimal
        lblReadout = Format(0, "0.")
    ElseIf DecimalFlag = False Then
        lblReadout.Caption = lblReadout.Caption '+ Decimal
    End If

    DecimalFlag = True              ' Decimal point here does indicate fraction.
    LastInput = "NUMS"
    pctKeyTrap.SetFocus
End Sub

Private Sub CmdMemoria_Click(Index As Integer)
    'Comandos de memoria
    Dim iMemoria
    Dim ValorMem As Double
    Dim sValorM As String * ANCHOMEM
    Dim temp$, i, b$, c$
    Dim Millar$

    Millar$ = ","
    If Decimal1 = "," Then
        Millar$ = "."
    End If
    '
    iMemoria = Val(LblMem.Caption)
    temp$ = Mid$(LstMem.List(iMemoria), 3)
    b$ = ""
    For i = 1 To Len(temp$)
        c$ = Mid$(temp$, i, 1)
        If c$ <> Millar$ Then
            If c$ = Decimal1 Then c$ = "."
            b$ = b$ + c$
        End If
    Next
    ValorMem = Val(b$)
    'quitar las comas...
    temp$ = lblReadout
    If Decimal1 = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
    
    Select Case Index
        Case 0      'Suma
            ValorMem = ValorMem + Val(temp$)
        Case 1      'Asigna directamente
            ValorMem = Val(temp$)
        Case 2      'Pone el valor
            lblReadout.Caption = Format(ValorMem, "#.####")
            LastInput = "NUMS"
        Case 3      'Borra la memoria
            ValorMem = 0
        Case 4      'Resta
            ValorMem = ValorMem - Val(temp$)
    End Select
    RSet sValorM = Str$(ValorMem)
    MemAddItem iMemoria, sValorM
    pctKeyTrap.SetFocus
End Sub

' Click event procedure for number (0-9) keys.
' Appends new number to the number in the display.
'
Private Sub cmdNumber_Click(Indx As Integer)
    If LastInput <> "NUMS" Then
        'lblReadOut.Caption = ""
        lblReadout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        lblReadout.Caption = lblReadout.Caption + cmdNumber(Indx).Caption
    Else
        lblReadout = Left(lblReadout, InStr(lblReadout, Format(0, ".")) - 1) + cmdNumber(Indx).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then
        lblReadout = "-" + lblReadout
    End If
    LastInput = "NUMS"
    pctKeyTrap.SetFocus
End Sub

Private Sub CmdOff_Click()
    'Guardar los datos de la tira...
    GuardarMemorias
    Unload Me
    'End
End Sub

' Click event procedure for operator (+, -, x, /, =) keys.
' If the immediately preceeding keypress was part of a
' number, increment NumOps. If one operand is present,
' set Op1. If two are present, set Op1 equal to the
' result of the operation on Op1 and the current
' input string, and display the result.
'
Private Sub cmdOperator_Click(Indx As Integer)
    Dim temp$

    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    'Else
    '    LblReadOut = Format(0, ".")
    End If

    ' If "," is being used as decimal separator, convert
    ' back to "." for numeric calculations.
    temp$ = lblReadout.Caption
    If Decimal1 = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
    If NumOps = 0 Then
        If cmdOperator(Indx).Caption = "-" And LastInput <> "NEG" Then
            lblReadout = "-" + lblReadout
            LastInput = "NEG"
        End If
    ElseIf NumOps = 1 Then
        Op1 = Val(temp$)
        If cmdOperator(Indx).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            lblReadout = "-"
            LastInput = "NEG"
        End If
        If OpFlag <> "=" Then
            If nList = 0 Then
                TiraAddItem " ", temp$
            Else
                TiraAddItem OpFlag, temp$
            End If
        End If
        If Indx > 4 Then
            Select Case Indx 'OpFlag
                Case 6 '"1/X"
                    If Op1 = 0 Then
                        lblReadout.Caption = ";ERROR"
                        LastInput = "OPS"
                        pctKeyTrap.SetFocus
                        Exit Sub
                    Else
                        Op1 = 1 / Op1
                    End If
                Case 5 '"+/-"
                    Op1 = -Op1
                Case 7 '"D"
                    Op1 = Fix(Op1)
            End Select
            temp$ = LTrim$(Str$(Op1))
            'If Decimal = "," And InStr(temp$, ".") Then Mid$(temp$, InStr(temp$, "."), 1) = ","
            lblReadout.Caption = Format(Val(temp$), "#.####")
            NumOps = 1
        End If
    ElseIf NumOps = 2 Then
        TiraAddItem OpFlag, temp$
        Op2 = Val(temp$)
        Select Case OpFlag
            Case "+"
                Op1 = Op1 + Op2
            Case "-"
                Op1 = Op1 - Op2
            Case "*"
                Op1 = Op1 * Op2
            Case "/"
                If Op2 = 0 Then
                    lblReadout.Caption = ";ERROR"
                    LastInput = "OPS"
                    pctKeyTrap.SetFocus
                    Exit Sub
                Else
                    Op1 = Op1 / Op2
                End If
            Case "="                   'Nunca se cumple esta opci¢n
                Op1 = Op2              'ya que NumOps=1
        End Select
        ' If "," is being used as decimal separator, convert
        ' "." back to "," for display.
        temp$ = LTrim$(Str$(Op1))
        If Indx = 4 Then
            TiraAddItem "=", temp$
        End If
        'If Decimal = "," And InStr(temp$, ".") Then Mid$(temp$, InStr(temp$, "."), 1) = ","
        lblReadout.Caption = Format(Val(temp$), "#.####")
        NumOps = 1
    End If
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = cmdOperator(Indx).Caption
    End If
    pctKeyTrap.SetFocus
End Sub

' Click event procedure for percent (%) key.
' Compute and display a percentage of the first operand.
'
Private Sub cmdPercent_Click()
    Dim temp$

    ' If "," is being used as decimal separator, convert
    ' back to "." for numeric calculations, then back to
    ' "," for display.
    temp$ = lblReadout.Caption
    If Decimal1 = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
    temp$ = LTrim$(Str$(Op1 * Val(temp$) / 100))
    'If Decimal = "," And InStr(temp$, ".") Then Mid$(temp$, InStr(temp$, "."), 1) = ","
    lblReadout.Caption = Format(Val(temp$), "#.####")

    pctKeyTrap.SetFocus
End Sub

' Initialization routine for the form.
' Set all variables to initial values.
'
Private Sub Form_Load()
    Dim temp$, i
    'Sólo una copia cada vez.
    If App.PrevInstance Then
        End
    End If
    
    frmH = Height
    frmW = Width
    ' Center form on screen.
    Top = (Screen.Height - Height) \ 2
    Left = (Screen.Width - Width) \ 2
    
    LstTira.Height = 2160
    LstMem.Height = LstTira.Height

    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = ""
    
    'Número de operaciones almacenadas
    MaxItemEnList = 200
    LblMem.Caption = "0"
    mnuBorMem_Click
    'Leer el fichero de memorias...
    LeerMemorias
    If frmL And frmT Then
        Left = frmL
        Top = frmT
    End If
    frmL = Left
    frmT = Top
    ' Determine whether "." or "," should be used as
    ' decimal separator based on value returned by
    ' FORMAT$ (country specific).
    temp$ = Format$(1.5, "#.#")
    If InStr(temp$, ",") Then
        Decimal1 = ","
        cmdDecimal.Caption = ","
    Else
        Decimal1 = "."
    End If
    lblReadout = Format(0, "0.")
    nPos = 1
    nPos1 = 1
End Sub

Private Sub Form_Resize()
    If WindowState <> 1 Then
        If frmH <> Height Then
            Height = frmH
        End If
        If frmW <> Width Then
            Width = frmW
        End If
    End If
End Sub

Private Sub GuardarMemorias()
    'Guardar el contenido de las memorias y
    'las últimas operaciones
    Dim nFic As Integer
    Dim sFic As String
    Dim i, j, b$

    frmL = Me.Left
    frmT = Me.Top
    sFic = "c:\Memorias.200"
    nFic = FreeFile
    Open sFic For Output As nFic
    Print #nFic, Str$(MaxItemEnList) + " ;Número de líneas en la tira de papel"
    Print #nFic, Str$(frmT) + " ;Posición superior  de la ventana"
    Print #nFic, Str$(frmL) + " ;Posición izquierda de la ventana"
    For i = 0 To 9
        b$ = LstMem.List(i)
        Print #nFic, b$
    Next
    j = LstTira.ListCount - 1
    For i = 0 To j
        b$ = LstTira.List(i)
        Print #nFic, b$
    Next
    Close
End Sub

Private Sub LeerMemorias()
    Dim nFic As Integer
    Dim sFic As String
    Dim i, j, b$

    mnuEdiTir_Click                    'Borrar el List
    sFic = "c:\Memorias.200"
    If Len(Dir$(sFic)) Then
        nFic = FreeFile
        Open sFic For Input As nFic
        Line Input #nFic, b$
        MaxItemEnList = Val(b$)
        Line Input #nFic, b$
        frmT = Val(b$)
        Line Input #nFic, b$
        frmL = Val(b$)
        For i = 0 To 9
            Line Input #nFic, b$
            LstMem.RemoveItem i
            LstMem.AddItem b$, i
        Next

        Do While Not EOF(nFic)
            Line Input #nFic, b$
            If Len(LTrim$(b$)) Then
                LstTira.AddItem b$
                nList = nList + 1
            End If
        Loop
        Close
    End If
    If nList Then
        EnList = True
        LstTira.ListIndex = nList - 1
        EnList = False
    End If

End Sub

Private Sub Label2_Click()

End Sub

Private Sub LblMenu_Click()
  PopupMenu mnuArc, , LblMenu.Left, LblMenu.Top + 250
End Sub

Private Sub LstMem_Click()
    'Asignar el valor del list al caption
    '-----------------------------------------------------
    'Quitar lo comentado, si se quiere que se asigne
    'a la pantalla el valor de la memoria pulsada
    '-----------------------------------------------------
    'Dim temp$, c$, b$
    Dim i
    'Dim Millar$

    If EnList Then Exit Sub

    'temp$ = LblReadOut.Caption
    'If Decimal = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
    'Op1 = Val(temp$)

    'Millar$ = ","
    'If Decimal = "," Then
    '    Millar$ = "."
    'End If
    
    i = LstMem.ListIndex
    LblMem = Str$(i)
    'If i >= 0 Then
    '    temp$ = Trim$(LstMem.List(i))
    '    i = InStr(temp$, " ")
    '    If i Then temp$ = Mid$(temp$, i + 1)
    '    For i = 1 To Len(temp$)
    '        c$ = Mid$(temp$, i, 1)
    '        If c$ <> Millar$ Then
    '            If c$ = Decimal Then c$ = "."
    '            If InStr("0123456789.-", c$) Then
    '                b$ = b$ + c$
    '            End If
    '        End If
    '    Next
    '    LblReadOut.Caption = Format(Val(b$), "#.####")
    '    LastInput = "NUMS"
    'End If
    pctKeyTrap.SetFocus
End Sub

Private Sub LstTira_Click()
    'Asignar el valor del list al caption

    Dim temp$, c$, b$
    Dim i
    Dim Millar$

    If EnList Then Exit Sub

    temp$ = lblReadout.Caption
    If Decimal1 = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
    Op1 = Val(temp$)

    Millar$ = ","
    If Decimal1 = "," Then
        Millar$ = "."
    End If
    
    i = LstTira.ListIndex
    If i >= 0 Then
        temp$ = Trim$(LstTira.List(i))
        i = InStr(temp$, " ")
        If i Then temp$ = Mid$(temp$, i + 1)
        For i = 1 To Len(temp$)
            c$ = Mid$(temp$, i, 1)
            If c$ <> Millar$ Then
                If c$ = Decimal1 Then c$ = "."
                If InStr("0123456789.-", c$) Then
                    b$ = b$ + c$
                End If
            End If
        Next
        lblReadout.Caption = Format(Val(b$), "#.####")
        LastInput = "NUMS"
    End If
    pctKeyTrap.SetFocus
End Sub

Private Sub MemAddItem(nMem, Valor$)
    'Añadir a la Lista de memorias
    Dim sEntero As String * ANCHOMEM
    Dim temp$

    temp$ = Format$(CDbl(Val(Valor$)), "###,###.00")
    LSet sEntero = Right$(Space$(ANCHOMEM) + temp$, ANCHOMEM)
    EnList = True
    LstMem.RemoveItem nMem
    LstMem.AddItem LTrim$(Str$(nMem)) + "=" + sEntero, nMem
    EnList = False
End Sub

Private Sub mnuAcerca_Click()
    Dim temp$
    temp$ = temp$ & "           Calculadora para Windows (PITICO) " & vbCrLf
    temp$ = temp$ & "    Modificada para Viveres y Licores S.A. de C.V." & vbCrLf & vbCrLf
    temp$ = temp$ & "       (con 10 memorias y lista de las 200 últimas operaciones)" & vbCrLf
    temp$ = temp$ & "           g32Calc v" & App.Major & "." & App.Minor & " (32 bits) (c) Guillermo Som Cerezo" & vbCrLf & vbCrLf
    temp$ = temp$ & "parte del código Copyright (C) 1982-1992 Microsoft Corporation    "
    MsgBox temp$, 0 + vbInformation, "Calculadora"
    pctKeyTrap.SetFocus
End Sub

Private Sub mnuArcGua_Click()
    GuardarMemorias

End Sub

Private Sub mnuArcLee_Click()
    LeerMemorias
End Sub

Private Sub mnuArcSalir_Click()
    CmdOff_Click
End Sub

Private Sub mnuBorMem_Click()
    'Borrar el contenido de las memorias
    Dim sValorM As String * ANCHOMEM
    Dim i, j
    
    j = LstMem.ListCount - 1
    For i = 0 To j
        LstMem.RemoveItem 0
    Next
    RSet sValorM = Format$(0, "#.00")
    'Ponerlas vacías
    For i = 0 To 9
        LstMem.AddItem LTrim$(Str$(i)) + "=" + sValorM
    Next
End Sub


Private Sub mnuEdiTir_Click()
    'Borrar el contenido de la Tira
    Dim i, j
    j = LstTira.ListCount - 1
    For i = 0 To j
        LstTira.RemoveItem 0
    Next
    nList = 0
End Sub

Private Sub mnuMinimizar_Click()
    WindowState = 1
End Sub


Private Sub MnuPrin_Click()
  mnuArc.Visible = True
End Sub

' KeyPress event procedure for KeyTrap picture box.
' The hidden KeyTrap control acts as the central keyboard
' handler. It processes keyboard input and invokes appropriate
' Event Procedure.  Focus is always returned to the
' KeyTrap picture box at the end of each click event
' for the command buttons.
'
Private Sub pctKeyTrap_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57:                                '0-9
            cmdNumber_Click KeyAscii - 48
        Case 8                                        'Borrar el último número
            CmdBack_Click
        Case 27:                                      'Cancel Entry (CE)
            cmdCancelEntry_Click
        Case 13:                                      'Enter is same as Equal (=)
            cmdOperator_Click 4
        Case Asc("C"), Asc("c")                       'Cancel All (C)
            cmdCancel_Click
        Case Asc("D"), Asc("d")                       'Cambia de signo (D)
            cmdOperator_Click 7
        Case 37:                                      'Percent (%)
            cmdPercent_Click
        Case 42:                                      'Multiply (*)
            cmdOperator_Click 2
        Case 43:                                      'Add (+)
            cmdOperator_Click 0
        Case 45:                                      'Subtract (-)
            cmdOperator_Click 1
        Case 44, 46:                                  'Decimal (.) or Comma (,)
            cmdDecimal_Click
        Case 47:                                      'Divide (/)
            cmdOperator_Click 3
        Case 61:                                      'Equal (=)
            cmdOperator_Click 4
        Case 97:
            LblMenu_Click
        Case Else: Beep                               'Reject all others
           
    End Select
End Sub

Private Sub pctKeyTrap_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then CmdOff_Click
If KeyCode = 27 Then CmdOff_Click
End Sub

Private Sub scrMem_Change()
    LblMem = LTrim$(Str$(scrMem.Value))
    pctKeyTrap.SetFocus
End Sub

Private Sub TiraAddItem(Oper$, Valor$)
    'A¤adir a la Lista el valor y la operaci¢n
    Dim temp$, c$
    Dim sEntero As String * ANCHOTIRA
    Dim sOper As String * 2

    'Si ya tenemos todos los apuntes
    'del papel, quitar el primero
    If nList >= MaxItemEnList Then
        LstTira.RemoveItem 0
        nList = nList - 1
    End If
    
    temp$ = Format$(CDbl(Val(Valor$)), "###,###.00")
    EnList = True
    If Oper$ = "1/X" Then
        c$ = "¼"    'Alt+0188
    ElseIf Oper$ = "+/-" Then
        c$ = "±"    'Alt+0177
    Else
        c$ = Oper$
    End If
    LSet sOper = c$
    LSet sEntero = Right$(Space$(ANCHOTIRA) + temp$, ANCHOTIRA)
    LstTira.AddItem sOper + sEntero
    LstTira.ListIndex = nList
    nList = nList + 1
    EnList = False
End Sub

Private Sub tim1_Timer()
Dim cMens As String
Dim n As Integer
cMens = "VIVERES Y LICORES S.A DE C.V. [PITICO]" + Space(3)
Me.lbleti.Caption = Mid(cMens, 1, nPos)
If lbleti.Width <= TextWidth(Mid(cMens, 1, nPos)) Then
   Me.lbleti.Caption = Mid(cMens, nPos1)
   nPos1 = nPos1 + 1
   If nPos1 = Len(cMens) Then
      nPos1 = 1: nPos = 1
   End If
End If
nPos = nPos + 1
End Sub
