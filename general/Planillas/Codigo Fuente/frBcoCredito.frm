VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frBcoCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Haberes: Banco de Crédito"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frBcoCredito.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmInfoExpress 
      Caption         =   "&Infoexpress"
      Height          =   360
      Left            =   4245
      TabIndex        =   19
      Top             =   2790
      Width           =   1425
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4245
      TabIndex        =   14
      Top             =   3945
      Width           =   1425
   End
   Begin VB.CommandButton cmdTransferir 
      Caption         =   "&Transferir"
      Height          =   360
      Left            =   4245
      TabIndex        =   13
      ToolTipText     =   "Tranferir a disco. Genera el archivo PagHab.Dat el cual se presentara al banco"
      Top             =   3150
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información de la transacción"
      Height          =   1605
      Left            =   90
      TabIndex        =   9
      Top             =   2655
      Width           =   3825
      Begin MSComCtl2.DTPicker xFecha 
         Height          =   300
         Left            =   2265
         TabIndex        =   8
         ToolTipText     =   "Fecha del abono a las cuentas"
         Top             =   1035
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24379393
         CurrentDate     =   36763
      End
      Begin VB.Label xImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   270
         Left            =   2265
         TabIndex        =   16
         ToolTipText     =   "Total a abonar. este monto se descontara del saldo de la cuenta corriente de la empresa"
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label xTotAbono 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   270
         Left            =   2265
         TabIndex        =   15
         ToolTipText     =   "Número de abonos a realizar"
         Top             =   435
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de pago"
         Height          =   195
         Left            =   375
         TabIndex        =   12
         ToolTipText     =   "Fecha del abono a las cuentas"
         Top             =   1065
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Importe a abonar"
         Height          =   195
         Left            =   375
         TabIndex        =   11
         ToolTipText     =   "Total a abonar. este monto se descontara del saldo de la cuenta corriente de la empresa"
         Top             =   750
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total de Abonos"
         Height          =   195
         Left            =   375
         TabIndex        =   10
         ToolTipText     =   "Número de abonos a realizar"
         Top             =   450
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalles de Cabezera"
      Height          =   2505
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   5640
      Begin AplisetControlText.Aplitext xNumCta 
         Height          =   285
         Left            =   2655
         TabIndex        =   7
         ToolTipText     =   "Número de la cuenta corriente a la que se le hará el cargo (Cuenta Cte. de la empresa)"
         Top             =   1935
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         Text            =   "0"
         Entero          =   -1  'True
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xSuc 
         Height          =   285
         Left            =   2655
         TabIndex        =   5
         ToolTipText     =   "Código de la Sucursal, plaza o ciudad a la que pertenece la cuenta bancaria de la empresa"
         Top             =   1575
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Text            =   "0"
         Entero          =   -1  'True
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xCodAfi 
         Height          =   285
         Left            =   2655
         TabIndex        =   3
         ToolTipText     =   "Asignado por un Funcionario de Negocios del Banco de Crédito"
         Top             =   1215
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Text            =   "0"
         Entero          =   -1  'True
         SinBlancos      =   -1  'True
      End
      Begin VB.Label xMoneda 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4695
         TabIndex        =   18
         ToolTipText     =   "Moneda de Abono de la transferencia"
         Top             =   1215
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   3990
         TabIndex        =   17
         Top             =   1260
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número de la Cuenta Corriente"
         Height          =   195
         Left            =   345
         TabIndex        =   6
         ToolTipText     =   "Número de la cuenta corriente a la que se le hará el cargo (Cuenta Cte. de la empresa)"
         Top             =   1995
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código de la Sucursal"
         Height          =   195
         Left            =   345
         TabIndex        =   4
         ToolTipText     =   "Código de la Sucursal, plaza o ciudad a la que pertenece la cuenta bancaria de la empresa"
         Top             =   1627
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código de Afiliación"
         Height          =   195
         Left            =   345
         TabIndex        =   2
         ToolTipText     =   "Asignado por un Funcionario de Negocios del Banco de Crédito"
         Top             =   1260
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "El Banco de Crédito no permite abonos en cuentas jurídicas. No se encuentran disponibles abonos en cuentas de moneda extranjera."
         Height          =   630
         Left            =   1005
         TabIndex        =   1
         Top             =   405
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   300
         Picture         =   "frBcoCredito.frx":044A
         Top             =   450
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Personalizado"
      Height          =   360
      Left            =   4245
      TabIndex        =   20
      Top             =   3525
      Width           =   1425
   End
End
Attribute VB_Name = "frBcoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMDCANCELAR_CLICK()
    Unload Me
End Sub
Private Sub CMDTRANSFERIR_Click()
    On Error GoTo Err1
            Dim xFile As String, CadBan As String, xCad As String
            frSelDir.Show 1
            If VPTAREA = "" Then Exit Sub
            If Right(VPTAREA, 1) <> "\" Then VPTAREA = VPTAREA & "\"
            xFile = VPTAREA & "PAGHAB.DAT"
            If Dir$(xFile) <> "" Then
                If MsgBox("YA EXISTE EN ESTA RUTA UN ARCHIVO CORRESPONDIENTE AL PAGO DE REMUNERACIONES, DESEA UD. REEMPLAZAR EL ARCHIVO POR EL NUEVO QUE ESTÁ PROCESANDO", vbYesNo + vbQuestion) = vbNo Then Exit Sub
                Kill xFile
            End If
            Dim RSAUX As New ADODB.Recordset, SumTodo As Long
            SumTodo = SUMANUM(xNumCta.Text)
            RSAUX.Open "SELECT * FROM  [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic, adLockReadOnly
            Do While Not RSAUX.EOF
                SumTodo = SumTodo + SUMANUM(SoloNumeros(RSAUX!CTABANCO))
                RSAUX.MoveNext
            Loop
            Open xFile For Append As #1
            xCad = Trim(xImporte.Caption)
            xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(Trim(xImporte.Caption), 2)
            CadBan = "1" & Format(xCodAfi.Text, "000000") & Format(xSuc.Text, "00") & Format(Val(xNumCta.Text), String(13, "0")) & Left(REGSISTEMA.EMPRESA & String(21, " "), 21) & Format(SumTodo, String(15, "0")) & IIf(xMoneda.Caption = "MN", "S/", "US") & Format(Val(xCad), String(15, "0")) & Format(xfecha.Day, "00") & Format(xfecha.Month, "00")
            Print #1, CadBan
            RSAUX.MoveFirst
            Do While Not RSAUX.EOF
                CadBan = ""
                xCad = Trim(Format(RSAUX!Neto, "0.00"))
                xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(Trim(RSAUX!Neto), 2)
                CadBan = IIf(InStr(RSAUX!CTABANCO, "M") <> 0, "6", IIf(InStr(RSAUX!CTABANCO, "*") <> 0, "2", "4")) & Format(xCodAfi.Text, "000000") & Format(Val(SoloNumeros(RSAUX!CTABANCO)), String(16, "0")) & Left(RSAUX!NOMBRES & String(36, " "), 36) & IIf(xMoneda.Caption = "MN", "S/", "US") & Format(Val(xCad), String(15, "0")) & "   "
                Print #1, CadBan
                RSAUX.MoveNext
            Loop
            Close #1
            Set RSAUX = Nothing
            MsgBox "PROCESO COMPLETADO. SE HA GENERADO EL ARCHIVO " & xFile, vbInformation
            Exit Sub
Err1:
            MsgBox ERR.Description
            Exit Sub
End Sub
Private Sub CMINFOEXPRESS_Click()
    On Error GoTo Err1
    Dim xFile As String, CadBan As String, xCad As String
    frSelDir.Show 1
    If VPTAREA = "" Then Exit Sub
    If Right(VPTAREA, 1) <> "\" Then VPTAREA = VPTAREA & "\"
    xFile = VPTAREA & "CREDITO.TXT"
    If Dir$(xFile) <> "" Then
        If MsgBox("YA EXISTE EN ESTA RUTA UN ARCHIVO CORRESPONDIENTE AL PAGO DE REMUNERACIONES DEL BANCO DE CREDITO, DESEA UD. REEMPLAZAR EL ARCHIVO POR EL NUEVO QUE ESTÁ PROCESANDO", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Kill xFile
    End If
    Dim RSAUX As New ADODB.Recordset, SumTodo As Long, SumaCab
    RSAUX.Open "SELECT * FROM  [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic, adLockReadOnly
    SumaCab = 0
    Do While Not RSAUX.EOF
        SumaCab = SumaCab + Val(Right(Trim("" & RSAUX!CTABANCO), 11))
        RSAUX.MoveNext
    Loop
    SumaCab = SumaCab + Val(xNumCta.Text)
    Debug.Print SumaCab
    Open xFile For Append As #1
    xCad = Trim(xImporte.Caption)
    xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(Trim(xImporte.Caption), 2)
    CadBan = "*1HC" & Format(xSuc.Text, "000") & Format(Val(xNumCta.Text), String(11, "0")) & "      S/" & Format(Val(xCad), String(15, "0")) & Format(xfecha.Value, "DDMMYYYY") & String(20, " ") & Format(SumaCab, String(15, "0")) & Format(Val(xTotAbono.Caption), "000000") & "1" & String(15, " ") & "0"
    Print #1, CadBan
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        CadBan = " 2"
        xCad = Trim(Format(RSAUX!Neto, "0.00"))
        xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(xCad, 2)
        CadBan = CadBan & IIf(InStr("" & RSAUX!CTABANCO, "M") <> 0, "M", "A") & Format(Val(SoloNumeros("" & RSAUX!CTABANCO)), String(14, "0")) & "      " & Left(RSAUX!NOMBRES & String(40, " "), 40) & "S/" & Format(Val(xCad), String(15, "0")) & String(40, " ") & "0"

        Print #1, CadBan
        RSAUX.MoveNext
    Loop
    Close #1
    Set RSAUX = Nothing
    MsgBox "PROCESO COMPLETADO. SE HA GENERADO EL ARCHIVO " & xFile, vbInformation
    Exit Sub
Err1:
    MsgBox ERR.Description
    Resume Next
    Exit Sub
End Sub

Private Sub Command1_Click()
On Error GoTo Err1
    Dim xFile As String, CadBan As String, xCad As String
    '*********************************
    Dim FlagValidarIDC As String   '**validar IDDC    0:si nodesea validar  1:si desea validar
    '*********************************
    frSelDir.Show 1
    If VPTAREA = "" Then Exit Sub
    If Right(VPTAREA, 1) <> "\" Then VPTAREA = VPTAREA & "\"
    xFile = VPTAREA & "CREDITO.TXT"
    If Dir$(xFile) <> "" Then
        If MsgBox("Ya Existe en esta Ruta un archivo correspondiente al pago de remuneraciones del Banco de Credito, Desea Ud. reemplazar el Archivo por el nuevo que está procesando", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Kill xFile
    End If
    
    If MsgBox("Desea validar IDC vs CUENTA?", vbYesNo + vbQuestion) = vbYes Then
        FlagValidarIDC = "1"
    Else
        FlagValidarIDC = "0"
    End If

    Dim RSAUX As New ADODB.Recordset, SumTodo As Long, SumaCab
    RSAUX.Open "SELECT * FROM [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic, adLockReadOnly
    SumaCab = 0
    Do While Not RSAUX.EOF
        SumaCab = SumaCab + Val(Right(Trim("" & RSAUX!CTABANCO), 11))
        RSAUX.MoveNext
    Loop
    SumaCab = SumaCab + Val(xNumCta.Text)
    Debug.Print SumaCab
    Open xFile For Append As #1
    xCad = Trim(xImporte.Caption)
    xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(Trim(xImporte.Caption), 2)
    CadBan = "#1HC" & Format(xSuc.Text, "000") & Format(Val(xNumCta.Text), String(11, "0")) & "      S/" & Format(Val(xCad), String(15, "0")) & Format(xfecha.Value, "DDMMYYYY") & String(20, " ") & Format(SumaCab, String(15, "0")) & Format(Val(xTotAbono.Caption), "000000") & "1" & String(15, " ") & "0"
    Print #1, CadBan
    Debug.Print CadBan
    RSAUX.MoveFirst
    Dim TIPCTA As String
    Dim NROCTA As String, SERIE As String, CTATEXT As String, MONX As String
    
    
    Do While Not RSAUX.EOF
        CadBan = " 2"
        xCad = Trim(Format(RSAUX!Neto, "0.00"))
        xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(xCad, 2)
        'XXCAMBIO
        TIPCTA = ""
        If InStr("" & RSAUX!CTABANCO, "M") <> 0 Then
          TIPCTA = "M"
          ElseIf InStr("" & RSAUX!CTABANCO, "A") <> 0 Then TIPCTA = "A"
          ElseIf InStr("" & RSAUX!CTABANCO, "C") <> 0 Then TIPCTA = "C"
        End If
        TIPCTA = ESNULO(TIPCTA, "A")
        
        'CadBan = CadBan & TIPCTA & Format(Val(SoloNumeros("" & RSAUX!CTABANCO)), String(14, "0")) & "      " & Left(RSAUX!NOMBRES & String(40, " "), 40) & "S/" & Format(Val(xCad), String(15, "0")) & String(40, " ") & "0"
         CadBan = CadBan & TIPCTA & Format(Val(SoloNumeros("" & RSAUX!CTABANCO)), String(14, "0")) & "      " & Left(RSAUX!NOMBRES & String(40, " "), 40) & IIf(Mid(RSAUX!CTABANCO, 1, 1) = "$", "US", "S/") & Format(Val(xCad), String(15, "0")) & String(40, " ") & "0" & Left("DNI" & String(3, " "), 3) & Left(RSAUX!DOCIDEN & String(12, " "), 12) & FlagValidarIDC
        
         
        Print #1, CadBan
        Debug.Print CadBan
        RSAUX.MoveNext
    Loop
    Close #1
    Set RSAUX = Nothing
    MsgBox "Proceso completado. se ha generado el Archivo " & xFile, vbInformation
    Exit Sub
Err1:
    MsgBox ERR.Description
    Resume Next
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim DBBAN As New ADODB.Recordset
    Screen.MousePointer = 11
    Dim xCad As String, TIP1 As Long
    xfecha.Value = Date
    DBBAN.Open "SELECT CTABANCO FROM  [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic, adLockReadOnly
    Do While DBBAN.EOF
        If InStr(DBBAN!CTABANCO, "$") <> 0 Then TIP1 = TIP1 + 1
        DBBAN.MoveNext
    Loop
    If TIP1 = DBBAN.RecordCount Then 'LAS CUENTAS SON TODAS EN SOLES
        xMoneda.Caption = "ME"
    Else
        If TIP1 = 0 Then    'LAS CUENTAS SON TODAS EN SOLES
            xMoneda.Caption = "MN"
        Else
            MsgBox "DENTRO DE LOS TRABAJADORES SE HAN ENCONTRADO TRABAJADORES CON CUENTAS CORRIENTES TANTO EN SOLES COMO EN DOLARES. ESTA FORMA DE PAGO SOLO ADMITE UN TIPO DE MONEDA EN CADA ARCHIVO DE ABONO. NO SE PODRÁ SEGUIR CON LA OPERACIÓN", vbCritical
            cmdTransferir.Enabled = False
        End If
    End If
    xTotAbono.Caption = DBBAN.RecordCount & " "
    DBBAN.Close
    DBBAN.Open "SELECT SUM(NETO) AS TOTAL FROM  [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic, adLockReadOnly
    If Not DBBAN.EOF Then
        xImporte.Caption = Format(DBBAN!TOTAL, "0.00 ")
    End If
    Set DBBAN = Nothing
    Screen.MousePointer = 1
End Sub
Public Function SUMANUM(ByVal CUENTA As String) As Long
    Dim X As Integer, SUM As Byte
    SUM = 0
    For X = 1 To Len(CUENTA)
        SUM = SUM + Val(Mid(CUENTA, X, 1))
    Next
    SUMANUM = SUM
End Function

