VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form FrmMant5ta 
   Caption         =   "Mantenimiento de Quinta"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   Icon            =   "FrmMant5ta.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10575
   Begin MSDataGridLib.DataGrid DgQuinta 
      Height          =   3405
      Left            =   150
      TabIndex        =   10
      Top             =   2205
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   6006
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   11927551
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "Mes"
         Caption         =   "Mes"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Total Percibido"
         Caption         =   "Total Percibido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Proyectado Fin Año"
         Caption         =   "Proy. Fin Año"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Total Renta Percibir"
         Caption         =   "T. Renta Percibir"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Renta Afecta"
         Caption         =   "Renta Afecta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Impuesto Anual"
         Caption         =   "Impt. Anual"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Monto Retener"
         Caption         =   "Mon. Retener"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Rentencion Anterior"
         Caption         =   "Ret. Anterior"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "CODTRAB"
         Caption         =   "CODTRAB"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "NOMBRES"
         Caption         =   "NOMBRES"
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
      BeginProperty Column11 
         DataField       =   "ANNO"
         Caption         =   "AÑO"
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
      BeginProperty Column12 
         DataField       =   "ACUMULADO"
         Caption         =   "ACUMULADO"
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
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            DividerStyle    =   6
            Button          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   8895
      TabIndex        =   9
      Top             =   1710
      Width           =   1515
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   7200
      TabIndex        =   8
      Top             =   1710
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   7185
      TabIndex        =   5
      Top             =   960
      Width           =   3210
      Begin MSComCtl2.DTPicker xfecha 
         Height          =   315
         Left            =   1500
         TabIndex        =   6
         Top             =   195
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   24510467
         CurrentDate     =   37039
      End
      Begin VB.Label Label1 
         Caption         =   "Año  :"
         Height          =   210
         Left            =   645
         TabIndex        =   7
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trabajador"
      Height          =   1155
      Left            =   135
      TabIndex        =   0
      Top             =   960
      Width           =   6840
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   300
         Left            =   1860
         TabIndex        =   1
         Top             =   300
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   529
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   0
         Left            =   870
         TabIndex        =   4
         Top             =   345
         Width           =   765
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Datos"
         Height          =   195
         Index           =   1
         Left            =   885
         TabIndex        =   3
         Top             =   645
         Width           =   420
      End
      Begin VB.Label xDatos 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1860
         TabIndex        =   2
         Top             =   615
         Width           =   4230
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   285
      Picture         =   "FrmMant5ta.frx":08CA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmMant5ta.frx":1194
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   900
      TabIndex        =   13
      Top             =   75
      Width           =   9345
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   900
      Left            =   120
      Top             =   15
      Width           =   10320
   End
   Begin VB.Label Label3 
      Caption         =   "Total Monto Retenidos :"
      Height          =   240
      Left            =   5925
      TabIndex        =   12
      Top             =   5820
      Width           =   1755
   End
   Begin VB.Label Lbtotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   330
      Left            =   7815
      TabIndex        =   11
      Top             =   5745
      Width           =   1200
   End
End
Attribute VB_Name = "FrmMant5ta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents RSQUINTA As ADODB.Recordset
Attribute RSQUINTA.VB_VarHelpID = -1
Dim FLAGMOV As Boolean
Private Sub CMDACEPTAR_CLICK()
  On Error GoTo ERRP
    RSQUINTA.UpdateBatch
    If MsgBox("Esta seguro que desea procesar los datos. " & Chr(13) & "si escoge Si se grabará la información y reempleazará los valores en la planilla haciendo un recalculo del monto neto percibido." & Chr(13) & "si es No grabará sin reemplazar los valores.", vbQuestion + vbYesNo + vbDefaultButton2) = vbOK Then
        Call ACTQUINTA_BOL
        Unload Me
        Exit Sub
    End If
    Unload Me
  Exit Sub
ERRP:
    Resume Next
End Sub
Private Sub ACTQUINTA_BOL()
'NO SQL
    Dim RSAUX As New ADODB.Recordset
    Dim PERIODO As String, CONCEPTO As String, X As Integer
    Dim INUMBOL As Long
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT MES,[MONTO RETENER] FROM HIST5TA WHERE ANNO='" & Trim(Str(Year(xfecha))) & "' AND CODTRAB='" & xTrab.Tag & "' ORDER BY MES", DBSYSTEM
    Do While Not RSAUX.EOF
        CONCEPTO = DevuelveValor("SELECT CODIGO FROM CONCEPTOS WHERE TRIM(FORMULA)='_QUINTACAT'", DBSYSTEM)
        PERIODO = Trim(Format(RSAUX!MES, "00")) & Trim(Format(Year(xfecha), "0000"))
        If ExisteTabla("BOL" & PERIODO) Then
            DBSYSTEM.Execute "UPDATE BOL" & PERIODO & " SET TOTING=(TOTING-RENTA5TA)+" & RSAUX("MONTO RETENER") & ",RENTA5TA=" & RSAUX("MONTO RETENER") & " WHERE CODTRAB='" & xTrab.Tag & "'", X
            If X = 1 Then
                INUMBOL = DevuelveValor("SELECT INUMBOL FROM BOL" & PERIODO & "  WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
                DBSYSTEM.Execute "UPDATE MOV" & PERIODO & " SET MONTO=" & RSAUX("MONTO RETENER") & " WHERE CONCEPTO='" & CONCEPTO & "' AND INUMBOL=" & INUMBOL
            End If
        End If
        RSAUX.MoveNext
    Loop
End Sub

Private Sub CMDCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub DGQUINTA_AFTERCOLEDIT(ByVal COLINDEX As Integer)
    RSQUINTA.Update
    DgQuinta.Refresh
    CmdAceptar.SetFocus
    Lbtotal.Caption = Format(MESACUM_ANT(13, "MONTO RETENER"), "0.00 ")
    DgQuinta.SetFocus
End Sub

Private Sub DGQUINTA_BUTTONCLICK(ByVal COLINDEX As Integer)
If RSQUINTA.EOF Or RSQUINTA.BOF Then Exit Sub
If COLINDEX <> 1 Then Exit Sub
On Error GoTo ERRQUINTA:
If IsNull(RSQUINTA!MES) Then
    MsgBox "Tiene que ingresar el número de mes"
    Exit Sub
End If
CmdAceptar.SetFocus
If ESNULO(RSQUINTA("TOTAL PERCIBIDO"), 0) < 0 Then
    MsgBox "Tiene que colocar una cantidad en total percibido mayor a cero"
    Exit Sub
End If
DgQuinta.SetFocus

If MsgBox("Calcular 5ta correspondiente a este Mes", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If
CmdAceptar.SetFocus
Screen.MousePointer = 11
Dim RSCFGQUINTA As New ADODB.Recordset
'VARIABLES DE LA QUINTA
Dim TOTALPERCIBIDO As Double, PROYECTADOFINAL As Double, TOTALRENTAPERCIBIR As Double
Dim RENTAAFECTA As Double, IMPUESTOANUAL As Double, ACUMULADO As Double, SALDO As Double
Dim MONTORETENER As Double, RETENCIONANTERIOR As Double
Dim UIT1 As Double, UIT2 As Double, VALUIT As Double
Dim UIT3 As Double, UIT4 As Double, Porc1 As Double, Porc2 As Double, Porc3 As Double
    RENTAAFECTA = 0
    IMPUESTOANUAL = 0
    ACUMULADO = 0
    SALDO = 0
    MONTORETENER = 0

'VARIABLES DE CONFIGURACION DE QUINTA
 Dim XNUMM As Integer, VINGMIN As Double, RENTADEDUC As Single, XVALMES As Integer
 Dim XVALOPCION As Double, MESACUMULADO As Integer, MINIMO As Double
    
'VARIABLES AUXILIARES
 Dim M1 As Double, M2 As Double, S1 As Double, S2 As Double, M3 As Double, M4 As Double
 Dim XC As Double
        
    XNUMM = RSQUINTA!MES
    TOTALPERCIBIDO = RSQUINTA("TOTAL PERCIBIDO")
    DgQuinta.SetFocus
    Set RSCFGQUINTA = New ADODB.Recordset
    RSCFGQUINTA.Open "CONFIG5TA", DBSYSTEM, adOpenStatic
    RENTADEDUC = RSCFGQUINTA!NUMUIT * RSCFGQUINTA!VALORUIT 'AHORA SON 7 UIT * 3000
    VALUIT = RSCFGQUINTA!VALORUIT
    UIT1 = RSCFGQUINTA!UIT1
    UIT2 = RSCFGQUINTA!UIT2
    UIT3 = RSCFGQUINTA!UIT3
    UIT4 = RSCFGQUINTA!UIT4
    Porc1 = RSCFGQUINTA!PORCENTAJE
    Porc2 = RSCFGQUINTA!PORCENTAJE2
    Porc3 = RSCFGQUINTA!PORCENTAJE3
    
    
    XVALMES = Val(RSCFGQUINTA("MES" & Format(XNUMM, "00"))) 'NUMERO QUE DIVIDIRA
    MINIMO = RSCFGQUINTA!VALORMIN
    MESACUMULADO = XXNUMMESACUM(XNUMM, RSCFGQUINTA)
    
    Dim L As Double 'RENTA NETA ANUAL
    Dim M_ACUM_ANT_DELMES As Double
    Dim T_QUINTA_ANT As Double
    Dim XMESVAR As Integer
    
    XVALOPCION = 12 - XNUMM + 1 'MES PROYECTADA
    XMESVAR = 2
    If XNUMM > 7 And XNUMM <= 12 Then XMESVAR = 1
    'If XNUMM = 12 Then XMESVAR = 0
    'MONTOS ACUMULADOS AFECTOS A QUINTA DE LOS MESES ANTERIORES
    M_ACUM_ANT_DELMES = MESACUM_ANT(XNUMM, "TOTAL PERCIBIDO")
    
    'RETENCION DE QUINTAS ACUMULADAS DE LOS MESES ANTERIORES DE ACUERDO A LA VARIABLE MES ACUMULADO
    T_QUINTA_ANT = MESACUM_ANT(XNUMM, "MONTO RETENER", XNUMM - MESACUMULADO)
    If Not ESNULO(GetValor("SELECT XSUEINT FROM TRABAJADORES WHERE CODTRAB='" & Trim(xTrab.Text) & "'", DBSYSTEM), False) Then
        M1 = (TOTALPERCIBIDO * XVALOPCION) + M_ACUM_ANT_DELMES + (TOTALPERCIBIDO * XMESVAR)
      Else
        M1 = (TOTALPERCIBIDO * XVALOPCION) + M_ACUM_ANT_DELMES
    End If
    'M1 es la Proyeccion del Monto
    'L es la Renta Basica
    L = M1 - RENTADEDUC
    RETENCIONANTERIOR = MESACUM_ANT(XNUMM, "MONTO RETENER", XNUMM - 1)
    TOTALRENTAPERCIBIR = M1
    RENTAAFECTA = (TOTALRENTAPERCIBIR) - RENTADEDUC
    
    'Calculando el 1er. Tope
    If M1 > RENTADEDUC And M1 <= (VALUIT * UIT1) Then
        M2 = ((L * (Porc1 / 100)) - T_QUINTA_ANT) / XVALMES
        IMPUESTOANUAL = (RENTAAFECTA) * (Porc1 / 100)
        ACUMULADO = T_QUINTA_ANT
        SALDO = IMPUESTOANUAL - ACUMULADO
        MONTORETENER = M2
    ElseIf M1 > RENTADEDUC And M1 > (VALUIT * UIT1) Then
    'Calculando el 2do. Tope
        M3 = (UIT2 * VALUIT) * (Porc1 / 100)
        If L - (UIT2 * VALUIT) > (VALUIT * UIT1) And L - (UIT2 * VALUIT) <= (UIT3 * VALUIT) Then
            M4 = (L - (UIT2 * VALUIT)) * (Porc2 / 100)
        ElseIf L - (UIT2 * VALUIT) > (UIT4 * VALUIT) Then
            M4 = (L - (UIT2 * VALUIT)) * (Porc3 / 100)
        Else
            M4 = (L - (UIT2 * VALUIT)) * (Porc1 / 100)
        End If
        M2 = ((M3 + M4) - T_QUINTA_ANT) / XVALMES
        IMPUESTOANUAL = (M3 + M4)
        ACUMULADO = T_QUINTA_ANT
        SALDO = IMPUESTOANUAL - ACUMULADO
        MONTORETENER = M2
    End If
    
    'Calculo Antiguo
'    If (L) > (54 * RSCFGQUINTA!VALORUIT) Then 'SI ES MAYOR A 54 UIT'S
'        M3 = (54 * RSCFGQUINTA!VALORUIT) * (RSCFGQUINTA!PORCENTAJE / 100)
'        M4 = (L - (54 * RSCFGQUINTA!VALORUIT)) * (RSCFGQUINTA!PORCENTAJE2 / 100)
'        M2 = ((M3 + M4) - T_QUINTA_ANT) / XVALMES
'        IMPUESTOANUAL = (M3 + M4)
'        ACUMULADO = T_QUINTA_ANT
'        SALDO = IMPUESTOANUAL - ACUMULADO
'        MONTORETENER = M2
'    ElseIf (L) > (RENTADEDUC) Or M1 > (RENTADEDUC) Then
'        'M3 = (54 * RSCFGQUINTA!VALORUIT) * (RSCFGQUINTA!PORCENTAJE / 100)
'        M2 = ((L * (RSCFGQUINTA!PORCENTAJE / 100)) - T_QUINTA_ANT) / XVALMES
'        IMPUESTOANUAL = (RENTAAFECTA) * (RSCFGQUINTA!PORCENTAJE / 100)
'        ACUMULADO = T_QUINTA_ANT
'        SALDO = IMPUESTOANUAL - ACUMULADO
'        MONTORETENER = M2
'    End If
    'ACTUALIZANDO EL GRID CON EL NUEVO CALCULO
    If XNUMM < 12 And MONTORETENER < 0 Then MONTORETENER = 0
    'EN EL UTLIMO MES EL IMPORTE DE QUINTA CATEGORIA SI PUEDE SER NEGATIVO POR ERROR DE CALCULO
    'DE LOS MESES ANTERIORES INCLUYENDO EL MISMO ESO SE CONSIDERA COMKO UN DALDO
    DgQuinta.Columns(2) = Round(M1, 2)
    DgQuinta.Columns(3) = Round(TOTALRENTAPERCIBIR, 2)
    DgQuinta.Columns("RENTA AFECTA") = Round(RENTAAFECTA, 2)
    DgQuinta.Columns(5) = Round(IMPUESTOANUAL, 2)
    DgQuinta.Columns("SALDO") = Round(SALDO, 2)
    DgQuinta.Columns(7) = Round(MONTORETENER, 2)
    DgQuinta.Columns(8) = Round(RETENCIONANTERIOR, 2)
    DgQuinta.Columns("ACUMULADO") = Round(T_QUINTA_ANT, 2)
    
    Lbtotal.Caption = Format(MESACUM_ANT(13, "MONTO RETENER"), "0.00 ")
    Screen.MousePointer = 1
    Exit Sub
ERRQUINTA:
    MsgBox "POSIBLE INVALIDA OPERACION DEL USUARIO " & ERR.Description, vbExclamation
End Sub
Private Function MESACUM_ANT(MES As Integer, CAMPO As String, Optional Mes1 As Integer) As Double
On Error GoTo ERRMES
Dim MONTO As Double, I As Integer
Dim RSAUX As New ADODB.Recordset
    MONTO = 0
'CON RECORDSET
    Set RSAUX = RSQUINTA.Clone(adLockReadOnly)
    RSAUX.Filter = "MES <" & MES
    If RSAUX.EOF Or RSAUX.BOF Then Exit Function
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        If IsMissing(Mes1) Then
            If RSAUX!MES < MES Then MONTO = MONTO + RSAUX(CAMPO)
         Else:
            If RSAUX!MES >= Mes1 And RSAUX!MES < MES Then MONTO = MONTO + ESNULO(RSAUX(CAMPO), 0)
        End If
       RSAUX.MoveNext
    Loop
    MESACUM_ANT = MONTO
    Exit Function
ERRMES:
    MESACUM_ANT = 0
End Function
Private Function XXNUMMESACUM(NUMM As Integer, RSAUX1 As ADODB.Recordset) As Integer
    Dim MESACUMULADO As Integer
    Select Case NUMM 'SIRVE PARA SACAR EL ACUMULADO DE MESES ANTERIORES
        Case 1
            MESACUMULADO = RSAUX1!ACUMULA01
        Case 2
            MESACUMULADO = RSAUX1!ACUMULA02
        Case 3
            MESACUMULADO = RSAUX1!ACUMULA03
        Case 4
            MESACUMULADO = RSAUX1!ACUMULA04
        Case 5
            MESACUMULADO = RSAUX1!ACUMULA05
        Case 6
            MESACUMULADO = RSAUX1!ACUMULA06
        Case 7
            MESACUMULADO = RSAUX1!ACUMULA07
        Case 8
            MESACUMULADO = RSAUX1!ACUMULA08
        Case 9
            MESACUMULADO = RSAUX1!ACUMULA09
        Case 10
            MESACUMULADO = RSAUX1!ACUMULA10
        Case 11
            MESACUMULADO = RSAUX1!ACUMULA11
        Case 12
            MESACUMULADO = RSAUX1!ACUMULA12
    End Select
    XXNUMMESACUM = MESACUMULADO
End Function

Private Sub Form_Load()
    Me.Height = 6660
    Me.Width = 10695
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSQUINTA = Nothing
End Sub

Private Sub RSQUINTA_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    On Error GoTo ERRMOV
    If Not FLAGMOV Then Exit Sub
        DgQuinta.Columns("CODTRAB").Value = xTrab.Tag
        DgQuinta.Columns("NOMBRES").Value = VGUTIL(2)
        DgQuinta.Columns(11).Value = Year(xfecha)
    Exit Sub
ERRMOV:
    Exit Sub
End Sub

Private Sub XTRAB_DBLCLICK()
    'NO SQL
    Dim RS1 As New ADODB.Recordset
    Dim X As Integer
    RS1.Open "VWTRABAJ", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RS1
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xTrab.Tag = VGUTIL(1)
        xDatos.Caption = RS1!CENTRO
        DgQuinta.Enabled = True
        FLAGMOV = False
        Set RSQUINTA = New ADODB.Recordset
        RSQUINTA.Open "SELECT HIST5TA.MES, HIST5TA.[TOTAL PERCIBIDO], HIST5TA.[PROYECTADO FIN AÑO], HIST5TA.[TOTAL RENTA PERCIBIR], HIST5TA.[RENTA AFECTA], HIST5TA.[IMPUESTO ANUAL], HIST5TA.SALDO, HIST5TA.[MONTO RETENER],HIST5TA.CODTRAB,HIST5TA.NOMBRES,HIST5TA.ANNO,HIST5TA.[RENTENCION ANTERIOR],ACUMULADO " & _
                      "FROM HIST5TA WHERE ANNO='" & Year(xfecha) & "' AND CODTRAB='" & xTrab.Tag & "' ORDER BY MES", DBSYSTEM, adOpenKeyset, adLockBatchOptimistic
        Set DgQuinta.DataSource = RSQUINTA
        Lbtotal.Caption = Format(MESACUM_ANT(13, "MONTO RETENER"), "0.00 ")
        FLAGMOV = True
        DgQuinta.Columns("CODTRAB").Width = 0
        DgQuinta.Columns("NOMBRES").Width = 0
        DgQuinta.Columns(11).Width = 0
     Else
       DgQuinta.Enabled = False
    End If
    Set RS1 = Nothing
End Sub


