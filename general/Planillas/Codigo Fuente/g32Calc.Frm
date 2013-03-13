VERSION 5.00
Begin VB.Form frmCalc 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora"
   ClientHeight    =   4950
   ClientLeft      =   3585
   ClientTop       =   1830
   ClientWidth     =   4185
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "g32Calc.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   4185
   Tag             =   "Calculadora del Sistema"
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "M-"
      Height          =   345
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   4110
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
      ToolTipText     =   "Memoria"
      Top             =   30
      Width           =   1755
   End
   Begin VB.CommandButton CmdOff 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "OFF"
      Height          =   435
      Left            =   1140
      TabIndex        =   33
      ToolTipText     =   "Cerrar"
      Top             =   2790
      Width           =   1005
   End
   Begin VB.VScrollBar scrMem 
      Height          =   465
      Left            =   570
      Max             =   9
      TabIndex        =   32
      Top             =   2340
      Width           =   255
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "D"
      Height          =   345
      Index           =   7
      Left            =   3570
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4500
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "="
      Height          =   345
      Index           =   4
      Left            =   2220
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4500
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      Height          =   345
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4500
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      Height          =   345
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4110
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "+/-"
      Height          =   345
      Index           =   5
      Left            =   3570
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4110
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "*"
      Height          =   345
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      Height          =   345
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3330
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "1/X"
      Height          =   345
      Index           =   6
      Left            =   3570
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3330
      Width           =   495
   End
   Begin VB.CommandButton cmdDecimal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "."
      Height          =   345
      Left            =   1680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4500
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      Height          =   345
      Index           =   0
      Left            =   1140
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4500
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
      Height          =   345
      Index           =   1
      Left            =   1140
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4110
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      Height          =   345
      Index           =   2
      Left            =   1680
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4110
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "3"
      Height          =   345
      Index           =   3
      Left            =   2220
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4110
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      Height          =   345
      Index           =   6
      Left            =   2220
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
      Height          =   345
      Index           =   5
      Left            =   1680
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
      Height          =   345
      Index           =   4
      Left            =   1140
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "7"
      Height          =   345
      Index           =   7
      Left            =   1140
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3330
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
      Height          =   345
      Index           =   8
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3330
      Width           =   495
   End
   Begin VB.CommandButton cmdNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "9"
      Height          =   345
      Index           =   9
      Left            =   2220
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3330
      Width           =   495
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "M+"
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   4500
      Width           =   705
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MS"
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   705
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MR"
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   3330
      Width           =   705
   End
   Begin VB.CommandButton CmdMemoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MC"
      Height          =   345
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   2940
      Width           =   705
   End
   Begin VB.CommandButton cmdPercent 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
      Height          =   345
      Left            =   3570
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "C"
      Height          =   435
      Left            =   3570
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Inicializar"
      Top             =   2790
      Width           =   495
   End
   Begin VB.CommandButton cmdCancelEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CE"
      Height          =   435
      Left            =   3000
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Limpiar el ingreso"
      Top             =   2790
      Width           =   495
   End
   Begin VB.CommandButton CmdBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "<-"
      Height          =   435
      Left            =   2220
      TabIndex        =   27
      ToolTipText     =   "Borrar"
      Top             =   2790
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
      Left            =   1875
      TabIndex        =   29
      ToolTipText     =   "Tira"
      Top             =   30
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
      Top             =   5000
      Width           =   2760
   End
   Begin VB.Label lblReadout 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
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
      Tag             =   "Calculadora del Sistema"
      ToolTipText     =   "LED"
      Top             =   2310
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
      ToolTipText     =   "Seleccionar memoria"
      Top             =   2340
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
      ToolTipText     =   "Memoria seleccionada"
      Top             =   2580
      Width           =   435
   End
   Begin VB.Menu mnuArc 
      Caption         =   "&Archivo"
      Begin VB.Menu CopyMem 
         Caption         =   "Copiar Valor"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuArcLee 
         Caption         =   "&Leer"
      End
      Begin VB.Menu mnuArcGua 
         Caption         =   "&Guardar"
         Shortcut        =   {F9}
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
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuArcSep3 
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
Option Explicit

Const ANCHOMEM = 14
Const ANCHOTIRA = 16

Dim Op1 As Double
Dim Op2 As Double
Dim DecimalFlag As Integer
Dim NumOps As Integer
Dim LastInput As String
Dim OpFlag As String
Dim Sigdecimal As String
Dim nList As Integer
Dim EnList As Integer
Dim MaxItemEnList As Integer
Dim frmT, frmL
Dim frmH, frmW
Dim RegAct As RegWin

Private Sub CmdBack_Click()
    Dim temp$
    Dim C$
    Dim i%

    temp$ = LTrim$(RTrim$(lblReadout.Caption))
    i% = Len(temp$)
    If i > 2 Then
        If Right$(temp$, 1) = Sigdecimal Then
            temp$ = Left$(temp$, i - 2) + Sigdecimal
        Else
            'Quitar el último caracter
            temp$ = Left$(temp$, i - 1)
        End If
        If Sigdecimal = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
        lblReadout.Caption = Format(Val(temp$), "#.####")
    Else
        DecimalFlag = False
        lblReadout.Caption = Format(0, "0.")
        LastInput = "OPS"
        Beep
    End If
End Sub

Private Sub cmdCancel_Click()
    lblReadout.Caption = Format(0, "0.")
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = ""
    pctKeyTrap.SetFocus
End Sub

Private Sub cmdCancelEntry_Click()
    lblReadout.Caption = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
    pctKeyTrap.SetFocus
End Sub

Private Sub cmdDecimal_Click()
    If LastInput = "NEG" Then
        lblReadout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        lblReadout = Format(0, "0.")
    ElseIf DecimalFlag = False Then
        lblReadout.Caption = lblReadout.Caption
    End If

    DecimalFlag = True
    LastInput = "NUMS"
    pctKeyTrap.SetFocus
End Sub

Private Sub CmdMemoria_Click(Index As Integer)
    'Comandos de memoria
    Dim iMemoria
    Dim ValorMem As Double
    Dim sValorM As String * ANCHOMEM
    Dim temp$, i, b$, C$
    Dim Millar$

    Millar$ = ","
    If Sigdecimal = "," Then
        Millar$ = "."
    End If
    '
    iMemoria = Val(LblMem.Caption)
    temp$ = Mid$(LstMem.List(iMemoria), 3)
    b$ = ""
    For i = 1 To Len(temp$)
        C$ = Mid$(temp$, i, 1)
        If C$ <> Millar$ Then
            If C$ = Sigdecimal Then C$ = "."
            b$ = b$ + C$
        End If
    Next
    ValorMem = Val(b$)
    'quitar las comas...
    temp$ = lblReadout
    If Sigdecimal = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
    
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

Private Sub cmdNumber_Click(Indx As Integer)
    If LastInput <> "NUMS" Then
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
End Sub

Private Sub cmdOperator_Click(Indx As Integer)
    Dim temp$

    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    temp$ = lblReadout.Caption
    If Sigdecimal = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
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
        temp$ = LTrim$(Str$(Op1))
        If Indx = 4 Then
            TiraAddItem "=", temp$
        End If
        lblReadout.Caption = Format(Val(temp$), "#.####")
        NumOps = 1
    End If
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = cmdOperator(Indx).Caption
    End If
    pctKeyTrap.SetFocus
End Sub

Private Sub cmdPercent_Click()
    Dim temp$
    temp$ = lblReadout.Caption
    If Sigdecimal = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
    temp$ = LTrim$(Str$(Op1 * Val(temp$) / 100))
    lblReadout.Caption = Format(Val(temp$), "#.####")

    pctKeyTrap.SetFocus
End Sub

Private Sub CopyMem_Click()
    'Clipboard.SetText = Val(lblReadout.Caption)
End Sub

Private Sub Form_Activate()
    ActivarTools RegAct
End Sub

Private Sub Form_Load()
    Dim temp$, i
'    Sólo una copia cada vez.
'    frmH = Height
'    frmW = Width
'    Top = (Screen.Height - Height) \ 2
'    Left = (Screen.Width - Width) \ 2
    LstTira.Height = 2160
    LstMem.Height = LstTira.Height
    
    With RegAct
        .Buscar = False
        .Editar = False
        .Eliminar = False
        .Filtrar = False
        .Imprimir = False
        .Nuevo = False
        .Preliminar = False
    End With

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
    temp$ = Format$(1.5, "#.#")
    If InStr(temp$, ",") Then
        Sigdecimal = ","
        cmdDecimal.Caption = ","
    Else
        Sigdecimal = "."
    End If
    lblReadout = Format(0, "0.")
End Sub

Private Sub GuardarMemorias()
    'Guardar el contenido de las memorias y
    'las últimas operaciones
    Dim nFic As Integer
    Dim sFic As String
    Dim i, j, b$

    frmL = Left
    frmT = Top
    sFic = "Memorias.200"
    nFic = FreeFile
    Open sFic For Output As nFic
    Print #nFic, Str$(MaxItemEnList) + " ;N£mero de l¡neas en la tira de papel"
    Print #nFic, Str$(frmT) + " ;Posici¢n superior  de la ventana"
    Print #nFic, Str$(frmL) + " ;Posici¢n izquierda de la ventana"
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
    sFic = "Memorias.200"
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

Private Sub LstMem_Click()
    'Asignar el valor del list al caption
    '-----------------------------------------------------
    'Quitar lo comentado, si se quiere que se asigne
    'a la pantalla el valor de la memoria pulsada
    '-----------------------------------------------------
    Dim i

    If EnList Then Exit Sub
   
    i = LstMem.ListIndex
    LblMem = Str$(i)
    pctKeyTrap.SetFocus
End Sub

Private Sub LstTira_Click()
    'Asignar el valor del list al caption

    Dim temp$, C$, b$
    Dim i
    Dim Millar$

    If EnList Then Exit Sub

    temp$ = lblReadout.Caption
    If Sigdecimal = "," And InStr(temp$, ",") Then Mid$(temp$, InStr(temp$, ","), 1) = "."
    Op1 = Val(temp$)

    Millar$ = ","
    If Sigdecimal = "," Then
        Millar$ = "."
    End If
    
    i = LstTira.ListIndex
    If i >= 0 Then
        temp$ = Trim$(LstTira.List(i))
        i = InStr(temp$, " ")
        If i Then temp$ = Mid$(temp$, i + 1)
        For i = 1 To Len(temp$)
            C$ = Mid$(temp$, i, 1)
            If C$ <> Millar$ Then
                If C$ = Sigdecimal Then C$ = "."
                If InStr("0123456789.-", C$) Then
                    b$ = b$ + C$
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

Private Sub pctKeyTrap_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 48 To 57:                                '0-9
            cmdNumber_Click KeyAscii - 48
        Case 8                                        'Borrar el último número
            CmdBack_Click
        Case 27:                                      'Cancel Entry (CE)
            cmdCancelEntry_Click
        Case 13:                                      'Enter es Igual (=)
            cmdOperator_Click 4
        Case Asc("C"), Asc("c")                       'Cancelar todo(C)
            cmdCancel_Click
        Case Asc("D"), Asc("d")                       'Cambia de signo (D)
            cmdOperator_Click 7
        Case 37:                                      'Porcentaje (%)
            cmdPercent_Click
        Case 42:                                      'Multiplicar (*)
            cmdOperator_Click 2
        Case 43:                                      'Sumar (+)
            cmdOperator_Click 0
        Case 45:                                      'Restar (-)
            cmdOperator_Click 1
        Case 44, 46:                                  'Sigdecimal (.) or Comma (,)
            cmdDecimal_Click
        Case 47:                                      'Divide (/)
            cmdOperator_Click 3
        Case 61:                                      'Igual (=)
            cmdOperator_Click 4
        Case Else: Beep                               'Por otras teclas
    End Select
End Sub

Private Sub scrMem_Change()
    LblMem = LTrim$(Str$(scrMem.Value))
    pctKeyTrap.SetFocus
End Sub

Private Sub TiraAddItem(Oper$, Valor$)
    'A¤adir a la Lista el valor y la operaci¢n
    Dim temp$, C$
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
        C$ = "¼"    'Alt+0188
    ElseIf Oper$ = "+/-" Then
        C$ = "±"    'Alt+0177
    Else
        C$ = Oper$
    End If
    LSet sOper = C$
    LSet sEntero = Right$(Space$(ANCHOTIRA) + temp$, ANCHOTIRA)
    LstTira.AddItem sOper + sEntero
    LstTira.ListIndex = nList
    nList = nList + 1
    EnList = False
End Sub


Public Sub Comandos(ByVal Cadena As String)
    Beep
End Sub
