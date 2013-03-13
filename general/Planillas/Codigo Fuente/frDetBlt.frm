VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frDetBlt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "c"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "frDetBlt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   150
      TabIndex        =   2
      Top             =   4860
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   3900
      TabIndex        =   0
      Top             =   4845
      Width           =   1320
   End
   Begin MSDataGridLib.DataGrid dgDetalle 
      Height          =   4665
      Left            =   75
      TabIndex        =   1
      Top             =   120
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8229
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frDetBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSDET As New ADODB.Recordset
Dim RSORIGINAL As ADODB.Recordset
Dim RSBOLMESANO As ADODB.Recordset
Dim FECHAx As String
Dim FECHAd As Date
Dim intINUMBOL As Integer
Dim intNUMBOL As Integer
Dim intCODNUMBOL As Integer
Dim CODTRABx As String
Dim dTotTotin As Double
Dim dTotTotEg As Double
Public Property Let FECHAPROCESO(ByVal NEWDAT As Date)
    FECHAd = NEWDAT
    FECHAx = Format(Month(NEWDAT), "00") & Year(NEWDAT)
End Property
Public Property Let BOLMEANO_INUMBOL(ByVal NEWINBO As Integer)
    intINUMBOL = NEWINBO
End Property
Public Property Let BOLMEANO_NUMBOL(ByVal NEWNUBO As Integer)
    intNUMBOL = NEWNUBO
End Property
Public Property Let BOLMEANO_CODNUMBOL(ByVal NEWCODBOL As Double)
    intCODNUMBOL = NEWCODBOL
End Property
Public Property Let BOLMEANO_CODTRAB(ByVal NEWCODTRAB As String)
    CODTRABx = NEWCODTRAB
End Property
Public Property Let BOLMEANO_TOTIN(ByVal NEWTOTIN As Double)
    dTotTotin = NEWTOTIN
End Property
Public Property Let BOLMEANO_TOTEG(ByVal NEWTOTeg As Double)
    dTotTotEg = NEWTOTeg
End Property
Public Property Let BOLMEANO_QYERYORIGINAL(ByVal NEWQUERY As String)
    Set RSORIGINAL = New ADODB.Recordset
    RSORIGINAL.Open NEWQUERY, DBSYSTEM, adOpenDynamic, adLockReadOnly
End Property


Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()

If MsgBox("Desea guardar los cambios?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
Screen.MousePointer = 11
    Call procesar
Screen.MousePointer = 1

End Sub
Private Sub dgDetalle_RowColChange(LASTROW As Variant, ByVal LASTCOL As Integer)
On Error GoTo handler
    Dim XL As String
    If SEPUEDEMODIFICAR(RSDET.Fields(0)) = True Then
       dgDetalle.AllowUpdate = True
    Else
       dgDetalle.AllowUpdate = False
    End If
    
Exit Sub
handler:

End Sub
Private Sub Form_Load()
    
    RSDET.Open " [##_TMPDETBLT" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenDynamic, adLockOptimistic
    Set dgDetalle.DataSource = RSDET
    With dgDetalle
        .Columns("VALOR").Alignment = dbgRight
        .Columns("VALOR").NumberFormat = "##,##0.00"
        .Caption = VPTAREA
    End With
    
Set RSORIGINAL = New ADODB.Recordset
RSORIGINAL.Open " [##_TMPDETBLT" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenDynamic, adLockReadOnly
'--------------------------
Set RSBOLMESANO = New ADODB.Recordset
RSBOLMESANO.Open "SELECT * FROM BOL" & FECHAx & "   WHERE INUMBOL=" & intINUMBOL, DBSYSTEM, adOpenDynamic, adLockOptimistic
Call SETEARGRID(Me.dgDetalle)
End Sub
Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSDET = Nothing
End Sub
Private Function SEPUEDEMODIFICAR(ByVal VALOR As String) As Boolean
On Error GoTo handler
    Dim CAD As String
    Dim RS As ADODB.Recordset
    
    Set RS = New ADODB.Recordset
    CAD = "SELECT * FROM CONCEPTOS WHERE CODIGO='" & VALOR & "'  AND ESESCRITO=1"
    RS.Open CAD, DBSYSTEM, adOpenDynamic, adLockReadOnly
    If Not RS.EOF Then
        If RS!TIPO = 0 Then
        Else
        End If
        SEPUEDEMODIFICAR = True
    Else
        SEPUEDEMODIFICAR = False
    End If
Exit Function
handler:
        SEPUEDEMODIFICAR = False
End Function
Private Sub procesar()
Dim TIPOx As String
RSDET.MoveFirst
RSORIGINAL.MoveFirst
Do While Not RSDET.EOF
    If ESESCRITO(RSDET.Fields(0)) <> 0 Then
        If TIPO(RSDET.Fields(0)) = 0 Then 'si es cero al ASIS2000
           Call MODIFICA_ASIS2000(RSDET.Fields(0))
           Call MODIFICA_MOVMESANO(RSDET.Fields(0))
        Else
           If TIPO(RSDET.Fields(0)) = 1 Then 'si es 1 al INGMOV2000,MOVMESANO
              Call MODIFICA_INGMOV2000(RSDET.Fields(0))
              Call MODIFICA_MOVMESANO(RSDET.Fields(0))
              Call MODIFICA_BOLMESANO
           Else
              If TIPO(RSDET.Fields(0)) = 2 Then 'si es 2 al INGMOV2000,MOVMESANO,BOLMESANO
                 Call MODIFICA_INGMOV2000(RSDET.Fields(0))
                 Call MODIFICA_MOVMESANO(RSDET.Fields(0))
                 Call MODIFICA_BOLMESANO
              End If
           End If
        End If
    End If
    RSORIGINAL.MoveNext
    RSDET.MoveNext
Loop
RSDET.MoveFirst
End Sub
Private Function TIPO(ByVal COD As String) As Integer
On Error GoTo handler
    Dim CAD As String
    Dim RS As ADODB.Recordset
    
    Set RS = New ADODB.Recordset
    CAD = "SELECT * FROM CONCEPTOS WHERE CODIGO='" & COD & "'"
    RS.Open CAD, DBSYSTEM, adOpenDynamic, adLockReadOnly
    TIPO = RS!TIPO
Exit Function
handler:
    TIPO = 99
End Function
Private Function ESESCRITO(ByVal COD As String) As Integer
On Error GoTo handler
    Dim CAD As String
    Dim RS As ADODB.Recordset
    
    Set RS = New ADODB.Recordset
    CAD = "SELECT * FROM CONCEPTOS WHERE CODIGO='" & COD & "'"
    RS.Open CAD, DBSYSTEM, adOpenDynamic, adLockReadOnly
    ESESCRITO = RS!ESESCRITO
Exit Function
handler:
    ESESCRITO = 0
End Function
Private Sub MODIFICA_BOLMESANO()
Dim STRCURSOR As String
Dim K As Integer
Dim RSCURSOR As ADODB.Recordset

STRCURSOR = "SELECT SUMAAFP,SUMASALUD,SUMAIES,SUMARENTA,SUMASCTR,SUMACTS,SUMAGRAT,SUMAVAC FROM CONCEPTOS WHERE CODIGO='" & RSDET.Fields(0) & "'"
Set RSCURSOR = New ADODB.Recordset
RSCURSOR.Open STRCURSOR, DBSYSTEM, adOpenDynamic, adLockReadOnly
If RSCURSOR.EOF Then Exit Sub 'SI NO HAY NADA SALIR DEL PROCESO
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
If TIPO(RSDET.Fields(0)) = 1 Then
'*********************************************************************************************
For K = 0 To 7
    Select Case CStr(RSCURSOR.Fields(K).Name)
        Case "SUMAAFP": If RSCURSOR.Fields(K) = True Then RSBOLMESANO!SUMAAFP = RSBOLMESANO!SUMAAFP - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
        Case "SUMASALUD": If RSCURSOR.Fields(K) = True Then RSBOLMESANO!SUMASALUD = RSBOLMESANO!SUMASALUD - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
        Case "SUMAIES": If RSCURSOR.Fields(K) = True Then RSBOLMESANO!SUMAIES = RSBOLMESANO!SUMAIES - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
        Case "SUMARENTA": If RSCURSOR.Fields(K) = True Then RSBOLMESANO!SUMARENTA = RSBOLMESANO!SUMARENTA - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
        Case "SUMASCTR": If RSCURSOR.Fields(K) = True Then RSBOLMESANO!SUMASCTR = RSBOLMESANO!SUMASCTR - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
        Case "SUMACTS": If RSCURSOR.Fields(K) = True Then RSBOLMESANO!SUMACTS = RSBOLMESANO!SUMACTS - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
        Case "SUMAGRAT": If RSCURSOR.Fields(K) = True Then RSBOLMESANO!SUMAGRAT = RSBOLMESANO!SUMAGRAT - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
        Case "SUMAVAC": If RSCURSOR.Fields(K) = True Then RSBOLMESANO!SUMAVAC = RSBOLMESANO!SUMAVAC - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
    End Select
Next K
RSBOLMESANO!TOTING = RSBOLMESANO!TOTING - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
'*********************************************************************************************
Else:
    RSBOLMESANO!TOTEGR = RSBOLMESANO!TOTEGR - RSORIGINAL.Fields(2) + RSDET.Fields(2): RSBOLMESANO.Update
End If

End Sub
Private Sub MODIFICA_MOVMESANO(ByVal CONCEP As String)
     'TERMINADO
     DBSYSTEM.Execute "UPDATE MOV" & FECHAx & "  SET MONTO=" & RSDET.Fields(2) & "   WHERE INUMBOL=" & intINUMBOL & " AND CONCEPTO='" & CONCEP & "'  AND CODNOMBOL=" & intCODNUMBOL & " "
End Sub
Private Sub MODIFICA_ASIS2000(ByVal CONCEP As String)
     'TERMINADO
    DBSYSTEM.Execute "UPDATE ASIS2000 SET VALOR=" & RSDET.Fields(2) & "   WHERE CODTRAB='" & CODTRABx & "' AND CONCEPTO='" & CONCEP & "'  AND  DIA=" & FechS(FECHAd, Sqlf) & ""
End Sub
Private Sub MODIFICA_INGMOV2000(ByVal CONCEP As String)
     'TERMINADO
    DBSYSTEM.Execute "UPDATE INGMOV2000 SET VALOR=" & RSDET.Fields(2) & "   WHERE CODTRAB='" & CODTRABx & "' AND CONCEPTO='" & CONCEP & "'  AND CODNOMBOL=" & intCODNUMBOL & " "
End Sub
Private Sub SETEARGRID(MiDrid As DataGrid)
    Dim K As Integer
    For K = 0 To MiDrid.Columns.Count - 1
        MiDrid.Columns(K).Width = 1500
    Next K
End Sub

