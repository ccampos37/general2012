VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frBancoContinental 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Banco Continental"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frBancoContinental.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3105
      TabIndex        =   13
      Top             =   3030
      Width           =   1425
   End
   Begin VB.CommandButton cmGenerar 
      Caption         =   "&Generar"
      Height          =   345
      Left            =   3105
      TabIndex        =   12
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totales"
      Height          =   1350
      Left            =   105
      TabIndex        =   7
      Top             =   2025
      Width           =   2910
      Begin VB.Label xTotalSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   270
         Left            =   1305
         TabIndex        =   11
         Top             =   780
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total en Soles"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label xTotalAbonos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   270
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total de Abonos"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   390
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Banco"
      Height          =   1830
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   4425
      Begin AplisetControlText.Aplitext xCodServicio 
         Height          =   285
         Left            =   2130
         TabIndex        =   6
         Top             =   1185
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xCodEmpresa 
         Height          =   285
         Left            =   2130
         TabIndex        =   4
         Top             =   772
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xCodBanco 
         Height          =   285
         Left            =   2130
         TabIndex        =   2
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código de &Servicio"
         Height          =   195
         Left            =   525
         TabIndex        =   5
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código de &Empresa"
         Height          =   195
         Left            =   525
         TabIndex        =   3
         Top             =   817
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del &Banco"
         Height          =   195
         Left            =   525
         TabIndex        =   1
         Top             =   405
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frBancoContinental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmCerrar_Click()
    Unload Me
End Sub
Private Sub CMGENERAR_Click()
    Dim xFile As String, CadBan As String, xCad As String
    frSelDir.Show 1
    If VPTAREA = "" Then Exit Sub
    If Right(VPTAREA, 1) <> "\" Then VPTAREA = VPTAREA & "\"
    xFile = VPTAREA & Right("00000000" & xCodEmpresa.Text, 8) & ".TXT"
    If Dir$(xFile) <> "" Then
        If MsgBox("Ya existe en esta ruta un archivo correspondiente Pagos por el Banco, desea reemplazarlo", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Kill xFile
    End If
    Dim RSAUX As New ADODB.Recordset, SumTodo As Long, STRCTA As String
    RSAUX.Open "SELECT * FROM  [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic, adLockReadOnly
    Open xFile For Append As #1
    xCad = Trim(xTotalSoles.Caption)
    xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(Trim(xTotalSoles.Caption), 2)
    CadBan = "038010011" + xCodEmpresa.Text + Format(xTotalAbonos.Caption, "0000000") + Right(String(15, "0") & xCad, 15) + String(22, "0") + Format(Date, "YYYYMMDD") + String(34, " ")
    Print #1, CadBan
    Do While Not RSAUX.EOF
        CadBan = ""
        xCad = Format(RSAUX!Neto, "000000000000.00")
        xCad = Left(xCad, 12) & Right(xCad, 2)
        CadBan = "0680" + Left("" & RSAUX!DOCIDEN & String(15, " "), 15) + "001"
        STRCTA = SoloNumeros("" & RSAUX!CTABANCO)
        If Left(STRCTA, 4) = "0011" Then
            CadBan = CadBan + Left(STRCTA + String(20, "0"), 20)
        Else
            CadBan = CadBan + Right(String(20, " ") + STRCTA, 20)
        End If
        CadBan = CadBan + xCad + Format(Date, "YYYYMMDD")
        Print #1, CadBan
        RSAUX.MoveNext
    Loop
    Close #1
    Set RSAUX = Nothing
    MsgBox "Proceso completado. se ha generado el archivo" & xFile, vbInformation
    Exit Sub
Err1:
            MsgBox ERR.Description
            Exit Sub
End Sub

