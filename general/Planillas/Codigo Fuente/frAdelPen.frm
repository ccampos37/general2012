VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frAdelPen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adelantos Pendientes de Descuento"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frAdelPen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7110
   Tag             =   "Panel de Adelantos Pendientes de Descontar"
   Begin MSDataGridLib.DataGrid xData 
      Height          =   4395
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   7752
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
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
      Caption         =   "Adelantos Pendientes de Descuento"
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image imSuma 
      Height          =   240
      Left            =   5235
      Picture         =   "frAdelPen.frx":030A
      Top             =   4635
      Width           =   240
   End
   Begin VB.Label xSuma 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   270
      Left            =   5550
      TabIndex        =   2
      Top             =   4620
      Width           =   1245
   End
   Begin VB.Label lTransfer 
      AutoSize        =   -1  'True
      Caption         =   "Transferir a Cuenta Corriente"
      Height          =   195
      Left            =   525
      TabIndex        =   1
      Top             =   4665
      Width           =   2025
   End
   Begin VB.Image imTransfer 
      Height          =   240
      Left            =   195
      Picture         =   "frAdelPen.frx":064C
      Top             =   4665
      Width           =   240
   End
End
Attribute VB_Name = "frAdelPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REGACT As RegWin
Dim WithEvents RSADEL As ADODB.Recordset
Attribute RSADEL.VB_VarHelpID = -1

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    Dim STROPEN As String
    Set RSADEL = New ADODB.Recordset
    STROPEN = "SELECT CODIGO, TRABAJADORES.CODTRAB, LTRIM(APEPAT) + ' ' + LTRIM(APEMAT) + ' ' + LTRIM(NOMBRE) AS NOMBRES, MES, ADEL.MONTO FROM TRABAJADORES, " & RegSistema.TablaAdel & " ADEL WHERE ADEL.CODTRAB=TRABAJADORES.CODTRAB AND NOMBOL=0 ORDER BY ADEL.CODTRAB"
    RSADEL.Open STROPEN, DbSystem, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSADEL
    FORMATEARDG
    With REGACT
        .Buscar = True
        .Editar = True
        .ELIMINAR = True
        .Filtrar = False
        .IMPRIMIR = True
        .NUEVO = True
        .Preliminar = True
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSADEL = Nothing
End Sub

Public Sub FORMATEARDG()
    With xData
        .Columns("CODIGO").Visible = False
        .Columns("MONTO").NumberFormat = "0.00"
        .Columns("MONTO").Alignment = dbgRight
        .Columns("MES").Width = 900
        .Columns("NOMBRES").Width = 2654.929
    End With
    IMSUMA_Click
End Sub

Public Sub COMANDOTOOLBAR(COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            frAdelantos.Show
        Case "EDITAR"
            If RSADEL.RecordCount = 0 Then Exit Sub
            frValor.Show 1
            If vpTarea <> "" And vpTarea <> "0" Then
                DbSystem.Execute "UPDATE " & RegSistema.TablaAdel & " SET MONTO=" & vpTarea & " WHERE CODIGO=" & RSADEL!CODIGO
                RSADEL.Requery
                FORMATEARDG
            End If
        Case "ELIMINAR"
            If RSADEL.RecordCount = 0 Then Exit Sub
            If MsgBox("REALMENTE DESEA ELIMINAR LOS REGISTROS SELECCIONADOS", vbYesNo + vbInformation) = vbNo Then Exit Sub
            Dim XBOOK
            For Each XBOOK In xData.SelBookmarks
                RSADEL.Bookmark = XBOOK
                DbSystem.Execute "DELETE FROM " & RegSistema.TablaAdel & " WHERE CODIGO=" & RSADEL!CODIGO
            Next
            RSADEL.Requery
            FORMATEARDG
    End Select
End Sub

Private Sub IMSUMA_Click()
    Dim RSSUM As New ADODB.Recordset
    RSSUM.Open "SELECT SUM(MONTO) AS TOTAL FROM " & RegSistema.TablaAdel & " WHERE NUMBOL=0", DbSystem, adOpenStatic
    If RSSUM.RecordCount = 0 Then xSuma.Caption = "0.00" Else xSuma.Caption = Format(RSSUM!TOTAL, "##,##0.00")
    Set RSSUM = Nothing
End Sub

Private Sub IMTRANSFER_Click()
    If RSADEL.RecordCount = 0 Then Exit Sub
    If MsgBox("REALMENTE DESEA PASAR EL ADELANTO A UN MOVIMIENTO DE CUENTA CORRIENTE", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Dim RSGRUPOS As New ADODB.Recordset
    RSGRUPOS.Open "SELECT CODGRUPO, NOMBRE FROM CTAGRUPO WHERE TIPO=2 ORDER BY NOMBRE", DbSystem, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSGRUPOS
    frmComun.Show 1
    If vgUtil(1) = "" Then Exit Sub
    Set RSGRUPOS = Nothing
    vpTarea = "NUEVO"
    vpTrasPrm = vgUtil(1)
    Load frMoviCta
    With frMoviCta
        .xCodTrab.Text = RSADEL!CODTRAB & " : " & RSADEL!NOMBRES
        .xCodTrab.Locked = True
        .xCodTrab.Tag = RSADEL!CODTRAB
        .xDesc.Text = "TRANSF. DE ADELANTO DE " & AMeses(Month(RSADEL!MES)) & " DE " & Year(RSADEL!MES)
        .xCapital.Text = Format(RSADEL!MONTO, "0.00")
        .xCapital.Locked = True
        .xMoneda.ListIndex = 0
        .xMoneda.Locked = True
        .xFechaIni.Value = Date
        .xInteres.Text = "0"
        .xMeses.Text = "1"
        .Show 1
    End With
    If vpTarea = "ACEPTÓ" Then
                DbSystem.Execute "DELETE FROM " & RegSistema.TablaAdel & " WHERE CODIGO=" & RSADEL!CODIGO
                RSADEL.Requery
                FORMATEARDG
                MsgBox "LA TRASNFERENCIA SE COMPLETÓ EXITOSAMENTE", vbInformation
    End If
End Sub

Private Sub LTRANSFER_Click()
    IMTRANSFER_Click
End Sub

Private Sub XDATA_HEADCLICK(ByVal COLINDEX As Integer)
    RSADEL.Sort = xData.Columns(COLINDEX).DataField
End Sub

