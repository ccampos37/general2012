VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frCuentasCtesProg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Corrientes Programadas"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frCuentasCtesProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   2490
      TabIndex        =   13
      Top             =   6255
      Width           =   1140
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1065
      TabIndex        =   12
      Top             =   6255
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Caption         =   "Programación de Débitos"
      Enabled         =   0   'False
      Height          =   3045
      Left            =   150
      TabIndex        =   9
      Top             =   3090
      Width           =   4425
      Begin VB.CommandButton cmAjustar 
         Caption         =   "&Ajustar"
         Height          =   315
         Left            =   3225
         TabIndex        =   14
         Top             =   2655
         Width           =   1065
      End
      Begin MSDataGridLib.DataGrid xData 
         Height          =   2280
         Left            =   165
         TabIndex        =   11
         Top             =   330
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   4022
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
               LCID            =   10250
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
               LCID            =   10250
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Definición"
      Height          =   2880
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4425
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   375
         Left            =   195
         TabIndex        =   18
         Top             =   525
         Width           =   4170
         Begin AplisetControlText.Aplitext xTrab 
            Height          =   285
            Left            =   60
            TabIndex        =   19
            Top             =   45
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   503
            Locked          =   -1  'True
            Text            =   ""
         End
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mitad de Perido"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   2535
         Width           =   2970
      End
      Begin AplisetControlText.Aplitext xDesc 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   1110
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   285
         Left            =   1770
         TabIndex        =   7
         Top             =   2130
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Format          =   25034753
         CurrentDate     =   36814
      End
      Begin VB.CommandButton xGenerar 
         Caption         =   "&Generar"
         Height          =   300
         Left            =   3315
         TabIndex        =   8
         Top             =   2130
         Width           =   885
      End
      Begin AplisetControlText.Aplitext xNumDebitos 
         Height          =   285
         Left            =   1770
         TabIndex        =   6
         Top             =   1830
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         MaxLength       =   8
         Text            =   "0"
         Entero          =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xCapital 
         Height          =   285
         Left            =   1770
         TabIndex        =   4
         Top             =   1530
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         MaxLength       =   8
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin VB.Label xClave 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3195
         TabIndex        =   16
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   255
         TabIndex        =   15
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2175
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número de Débitos"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1875
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capital Total"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1575
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   315
         Width           =   765
      End
   End
End
Attribute VB_Name = "frCuentasCtesProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RSPROG As ADODB.Recordset
Attribute RSPROG.VB_VarHelpID = -1
Public MANT As Integer
Private Sub CMACEPTAR_CLICK()
    Dim xSuma As Single
    xSuma = DevuelveValor("SELECT SUM(IMPORTE) AS SUMA FROM  [##TMPCTACTEPROG" & VGL_COMPUTER & "] ", DBSYSTEM)
    If xSuma <> Valc(xCapital.Text) Then
        MsgBox "NO SE HA EJECUTADO LA OPCIÓN AJUSTAR", vbInformation
        Exit Sub
    End If
    DBSYSTEM.Execute "DELETE FROM CTACTEPROG WHERE CODMOV=" & VPCODTMP & "  AND " & _
          "CODTRAB ='" & VPNUMTMP & "'"
    If MANT = 1 Then
        xTrab.Tag = VPNUMTMP
        xSuma = VPCODTMP
        DBSYSTEM.Execute "UPDATE MOVICTA SET CAPITAL=" & Format(xCapital.Text, "0.00") & ",FECHAINI=" & FechS(xFechaIni.Value, Sqlf) & ",NUMMESES=" & _
                                                                 xNumDebitos.Text & ",SALDO=" & Format(xCapital.Text, "0.00") & ",CUOTA=0" & ",DESCRIPCION='" & xDesc.Text & "',PROGRAMADO=1,TIPOGRUPO=" & frCuentas.xTipo.ListIndex + 1 & _
                         " WHERE CODMOV=" & xSuma
     
        Else
        DBSYSTEM.Execute "INSERT INTO MOVICTA (CODTRAB,CODGRUPO,CAPITAL,FECHAINI,NUMMESES,SALDO, CUOTA,DESCRIPCION, PROGRAMADO, TIPOGRUPO) " & _
        "VALUES ('" & xTrab.Tag & "','" & VPTRASPRM & "'," & xCapital.Text & "," & DateSQL(xFechaIni.Value) & "," & xNumDebitos.Text & "," & xCapital.Text & ",0,'" & xDesc.Text & "',1," & frCuentas.xTipo.ListIndex + 1 & ")"
    End If
    
    RSPROG.MoveFirst
    If MANT = 0 Then
        xSuma = DevuelveValor("SELECT MAX(CODMOV) AS MCTA FROM MOVICTA", DBSYSTEM)
     ElseIf MANT = 1 Then xSuma = VPCODTMP
    End If
    Do While Not RSPROG.EOF
        DBSYSTEM.Execute "INSERT INTO CTACTEPROG (CODMOV,CODTRAB,FECHA,IMPORTE, SECUENCIA,MITAPER) " & _
        " VALUES (" & xSuma & ",'" & xTrab.Tag & "'," & DateSQL(RSPROG!FECHAINI) & "," & RSPROG!Importe & "," & RSPROG!SECUENCIA & ",'" & RSPROG!MITAD & "')"
        RSPROG.MoveNext
    Loop
    Call ACTSALDO(VPCODTMP)
    MsgBox "LA INFORMACIÓN FUE GRABADA SATISFACTORIAMENTE", vbInformation
    Unload Me
End Sub

Private Sub CMAJUSTAR_Click()
    Dim xSuma As Currency
    xSuma = Round(DevuelveValor("SELECT SUM(IMPORTE) AS SUMA FROM  [##TMPCTACTEPROG" & VGL_COMPUTER & "] ", DBSYSTEM), 2)
    If xSuma <> Valc(xCapital.Text) Then
        DBSYSTEM.Execute "UPDATE  [##TMPCTACTEPROG" & VGL_COMPUTER & "]  SET IMPORTE=IMPORTE+" & Valc(xCapital.Text) - xSuma & " WHERE SECUENCIA=" & xNumDebitos.Text
    End If
    REFRESCAR
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Frame3.Enabled Then Me.xTrab.SetFocus
End Sub

Private Sub Form_Load()
    Dim CAD As String
    Call CREARCAMPOS
    If ExisteTablaAux(" [##TMPCTACTEPROG" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTACTEPROG" & VGL_COMPUTER & "] "
    
    If MANT = 0 Then
        VPCODTMP = "''"
        VPNUMTMP = "''"
    End If
    CAD = "SELECT SECUENCIA=CAST(SECUENCIA AS INT),FECHAINI=CAST(FECHA AS DATETIME), " & _
          "IMPORTE=CAST(IMPORTE AS  Numeric(20,2) ),MITAD=CAST(ISNULL(MITAPER,'') AS CHAR(1)) INTO  [##TMPCTACTEPROG" & VGL_COMPUTER & "]  " & _
          "From CTACTEPROG WHERE CODMOV=" & VPCODTMP & "  AND " & _
          "CODTRAB ='" & VPNUMTMP & "' ORDER BY SECUENCIA ASC"
    DBSYSTEM.Execute CAD
    Set RSPROG = New ADODB.Recordset
    If MANT = 2 Then
        Frame1.Enabled = False
        RSPROG.Open " [##TMPCTACTEPROG" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockReadOnly
      Else
        RSPROG.Open " [##TMPCTACTEPROG" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenDynamic, adLockOptimistic
        cmAceptar.Enabled = True
    End If
    xDesc.Text = "" & VPTAREA
    REFRESCAR
    If MANT = 1 Or MANT = 2 Then
        Dim RSMOVI As New ADODB.Recordset
        RSMOVI.Open "SELECT * FROM MOVICTA WHERE CODMOV=" & VPCODTMP & "  AND " & _
          "CODTRAB ='" & VPNUMTMP & "'", DBSYSTEM, adOpenKeyset, adLockReadOnly
        xFechaIni.Value = RSMOVI!FECHAINI
        xCapital.Text = Format(RSMOVI!CAPITAL, "###,###,##0.00")
        xNumDebitos.Text = RSMOVI!NUMMESES
        Frame3.Enabled = False
    End If
    
End Sub

Private Sub xData_BeforeColEdit(ByVal COLINDEX As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    Dim RsTmp As New ADODB.Recordset
    Set RsTmp = New ADODB.Recordset
    RsTmp.Open "SELECT * FROM PAGOSCTA WHERE CODMOV=" & VPCODTMP & " AND CODTRAB='" & Trim(VPNUMTMP) & "' AND SECUENCIA=" & RSPROG("SECUENCIA"), DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RsTmp.RecordCount = 1 Then
        CANCEL = True
    End If
End Sub

Private Sub xData_BeforeColUpdate(ByVal COLINDEX As Integer, OldValue As Variant, CANCEL As Integer)
    Dim RsTmp As New ADODB.Recordset
    Set RsTmp = New ADODB.Recordset
    RsTmp.Open "SELECT * FROM PAGOSCTA WHERE CODMOV=" & VPCODTMP & " AND CODTRAB='" & Trim(VPNUMTMP) & "' AND SECUENCIA=" & RSPROG("Secuencia"), DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RsTmp.RecordCount = 1 Then
        CANCEL = True
    End If
End Sub

Private Sub XGENERAR_Click()
    Dim X As Integer, F As Date
    X = Valc(xNumDebitos.Text)
    If X < 1 Then
        MsgBox "ESTA OPCIÓN SOLO ES PERMITIDA PARA CUENTAS CORRIENTES PROGRAMADAS MAYORES O IGUALES A 2 DÉBITOS", vbInformation
        Exit Sub
    End If
    If Valc(xCapital.Text) <= 0 Then
        MsgBox "EL CAPITAL NO PUEDE SER 0 (CERO) O INFERIOR A ESTE VALOR", vbInformation
        Exit Sub
    End If
    F = xFechaIni.Value
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTACTEPROG" & VGL_COMPUTER & "]  "
    
    If Not IsNumeric(VPCODTMP) Then
        VPCODTMP = -1
    End If
    
    'CALCULANDO LA SUMA DE LAS CUENTAS CORRIENTES PROGRAMADAS Y DEBITADAS Y LA MAXIMA SECUENCIA
    Dim SqlCad As String
    SqlCad = "SELECT MAX(CTACTEPROG.SECUENCIA) AS MAXIMO,SUM(CTACTEPROG.IMPORTE) AS SUMIMP,MAX(FECHA) AS MAXFECHA " & _
             "FROM CTACTEPROG, PAGOSCTA " & _
             "WHERE " & _
             "(CTACTEPROG.SECUENCIA = PAGOSCTA.SECUENCIA) AND (CTACTEPROG.CODMOV = PAGOSCTA.CODMOV) AND  " & _
             "PAGOSCTA.CODMOV=" & VPCODTMP & " AND PAGOSCTA.CODTRAB='" & Trim(VPNUMTMP) & "'"
    Dim RSAUX As New ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open SqlCad, DBSYSTEM, adOpenKeyset, adLockReadOnly
    F = IIf(IsNull(RSAUX("MAXFECHA")), xFechaIni.Value, DateAdd("M", 1, RSAUX("MAXFECHA")))
    
    'INSERTANDO LAS FILAS YA COBRADAS
    SqlCad = "SELECT CTACTEPROG.SECUENCIA, CTACTEPROG.FECHA AS FECHAINI, CTACTEPROG.IMPORTE, CTACTEPROG.MITAPER AS MITAD " & _
             "FROM CTACTEPROG, PAGOSCTA " & _
             "WHERE " & _
             "(CTACTEPROG.SECUENCIA = PAGOSCTA.SECUENCIA) AND (CTACTEPROG.CODMOV = PAGOSCTA.CODMOV) AND  " & _
             "PAGOSCTA.CODMOV=" & VPCODTMP & " AND PAGOSCTA.CODTRAB='" & Trim(VPNUMTMP) & "'"
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTACTEPROG" & VGL_COMPUTER & "]  " & SqlCad
    
    For X = ESNULO(RSAUX("MAXIMO") + 1, 1) To Valc(xNumDebitos.Text)
        DBSYSTEM.Execute "INSERT INTO  [##TMPCTACTEPROG" & VGL_COMPUTER & "]  VALUES (" & X & "," & DateSQL(F) & "," & Round((Valc(xCapital.Text) - ESNULO(RSAUX("SUMIMP"), 0)) / (Valc(xNumDebitos.Text) - ESNULO(RSAUX("MAXIMO"), 0)), 2) & _
        IIf(Check1.Value = 0, ",''", ",'V'") & ")"
        F = DateAdd("M", 1, F)
    Next
    REFRESCAR
    cmAceptar.Enabled = True
    Frame2.Enabled = True
End Sub

Private Sub XTRAB_DBLCLICK()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT CODTRAB, NOMBRES, FECHAING, CODCCOSTO,CENTRO FROM VWTRABAJ WHERE SITUACIÓN<'2' AND CODTRAB NOT IN (SELECT CODTRAB FROM HISTOVAC WHERE CERRADO=0)", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSAUX.RecordCount = 0 Or RSAUX.EOF Then
        MsgBox "NO SE HAN ENCONTRADO TRABAJADORES", vbInformation
        Set RSAUX = Nothing
        cmAceptar.Enabled = False
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Text = RSAUX!NOMBRES
        xTrab.Tag = RSAUX!CODTRAB
    End If
    Set RSAUX = Nothing
End Sub

Public Sub REFRESCAR()
    With xData
        RSPROG.Requery
        Set .DataSource = RSPROG
        .Columns("IMPORTE").NumberFormat = "0.00 "
        .Columns("IMPORTE").Alignment = dbgRight
        .Columns("SECUENCIA").Locked = True
        .Columns("MITAD").Alignment = dbgCenter
    End With
End Sub
Private Sub CREARCAMPOS()
    If Not ExisteCampo("MITAPER", "CTACTEPROG", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE CTACTEPROG ADD MITAPER CHAR(1) NULL DEFAULT '' "
    End If
End Sub

