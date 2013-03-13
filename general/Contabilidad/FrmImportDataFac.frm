VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmImportDataFac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de datos y Generacion de Asiento Factura"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSerie 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   2025
      Width           =   3825
   End
   Begin VB.CheckBox ChkRest 
      Caption         =   "Con Restauracion"
      Height          =   240
      Left            =   60
      TabIndex        =   10
      Top             =   1155
      Width           =   2865
   End
   Begin MSComDlg.CommonDialog cmdg_archivo 
      Left            =   6570
      Top             =   3075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3540
      TabIndex        =   7
      Top             =   3075
      Width           =   1605
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   1920
      TabIndex        =   6
      Top             =   3075
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo de Proceso"
      Height          =   990
      Left            =   3420
      TabIndex        =   5
      Top             =   90
      Width           =   3510
      Begin MSComCtl2.DTPicker DTPPerido 
         Height          =   315
         Left            =   1005
         TabIndex        =   8
         Top             =   420
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM - yyyy"
         Format          =   51773443
         CurrentDate     =   37656
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo :"
         Height          =   270
         Left            =   150
         TabIndex        =   9
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.CommandButton CmdArchivo 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   6630
      TabIndex        =   4
      Top             =   2580
      Width           =   375
   End
   Begin TextFer.TxFer TxArchivo 
      Height          =   360
      Left            =   45
      TabIndex        =   3
      Top             =   2550
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   635
      BackColor       =   14546937
      Object.CausesValidation=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Text            =   ""
      ColorIlumina    =   14546937
      Valor           =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Generacion"
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   2835
      Begin VB.OptionButton Opt 
         Caption         =   "Servidor Local"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   645
         Width           =   2520
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Puntos Remotos"
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   300
         Width           =   2520
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccionar Punto Remoto"
      Height          =   240
      Left            =   75
      TabIndex        =   12
      Top             =   1770
      Width           =   3735
   End
End
Attribute VB_Name = "FrmImportDataFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim sqlcad As String
Dim cnxtrans As ADODB.Connection
Dim rstrans As ADODB.Recordset
Dim NombreArchivo As String
Dim SerieSubAsiento(10) As String
Private Sub CmdArchivo_Click()
    cmdg_archivo.Filter = "Archivos de Exportacion|EXPO*.EX"
    cmdg_archivo.ShowOpen
    NombreArchivo = cmdg_archivo.FileName
    TxArchivo.Text = NombreArchivo
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub CmdProcesar_Click()
Dim rsparimpo As ADODB.Recordset
Dim BaseOrigen As String
Dim paso1 As Integer
Dim procedimento As String
On Error GoTo Proceso
Screen.MousePointer = 11

paso1 = 1

Set rsparimpo = New ADODB.Recordset
Set rsparimpo = VGCNx.Execute(" select * from ct_importarventas")
If Opt(0).Value = True Then
    If ChkRest.Value = 1 Then
        If Not Restaurar Then
            Screen.MousePointer = 1
            Exit Sub
        End If
    End If
    BaseOrigen = "TRANSFER" & Trim$(Mid$(cboSerie.List(cboSerie.ListIndex), 7, 3))
End If
If Opt(1).Value = True Then
    BaseOrigen = rsparimpo!BaseVenta
End If
paso1 = rsparimpo!asientofacturacion
procedimiento = ESNULO(rsparimpo!procedimientoasiento, "")
Set COMANDO = New ADODB.Command
    Screen.MousePointer = 11
    '@BaseConta, @BaseVenta, @Ano, @Mes, @tipanal, @User
    VGGeneral.BeginTrans
    With COMANDO
        .CommandType = adCmdStoredProc
        .CommandText = "vt_insertacliente"
        .ActiveConnection = VGGeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@BaseVenta") = BaseOrigen
        .Parameters("@Mes") = Format(Month(DTPPerido.Value), "00")
        .Parameters("@Ano") = Year(DTPPerido.Value)
        .Parameters("@tipanal") = rsparimpo!tipanal
        .Parameters("@User") = VGParamSistem.Usuario
        .Execute
    End With
Set COMANDO = New ADODB.Command
If paso1 = 1 Then
   If procedimiento = "" Then procedimiento = "vt_generaasiento1_pro"
   With COMANDO
        .CommandType = adCmdStoredProc
        .CommandText = "" & procedimiento & ""  ' "vt_generaasiento1_pro"
        .ActiveConnection = VGGeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@BaseVenta") = VGParamSistem.BDEmpresa
        .Parameters("@BaseParam") = rsparimpo!BaseVenta
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@Libro") = rsparimpo!Libro
        .Parameters("@Mes") = Format(Month(DTPPerido), "00")
        .Parameters("@Ano") = Year(DTPPerido)
        .Parameters("@ctasoles") = rsparimpo!cuentasoles
        .Parameters("@ctadolares") = rsparimpo!cuentadolares
        .Parameters("@ctaIGV") = rsparimpo!ctaigv
        .Parameters("@tipanal") = rsparimpo!tipanal
        .Parameters("@Compu") = VGComputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Execute
    End With
Else
 If procedimiento = "" Then procedimiento = "vt_generaasiento_pro"
 With COMANDO
        .CommandType = adCmdStoredProc
        .CommandText = "" & procedimiento & ""     '  "vt_generaasiento_pro"
        .ActiveConnection = VGGeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@BaseVenta") = BaseOrigen
        .Parameters("@BaseParam") = rsparimpo!BaseVenta
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@Libro") = rsparimpo!Libro
        .Parameters("@Mes") = Format(Month(DTPPerido), "00")
        .Parameters("@Ano") = Year(DTPPerido)
        .Parameters("@ctatotal") = rsparimpo!ctatotal
'        .Parameters("@ctasoles") = rsparimpo!cuentasoles
'        .Parameters("@ctadolares") = rsparimpo!cuentadolares
        .Parameters("@ctaIGV") = rsparimpo!ctaigv
        .Parameters("@tipanal") = rsparimpo!tipanal
        .Parameters("@Compu") = VGComputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Execute
    End With
End If
    VGGeneral.CommitTrans
    Screen.MousePointer = 1
    MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Ventas"
    paso = 0
    Exit Sub
Proceso:
    Screen.MousePointer = 1
    MsgBox err.Description
    VGGeneral.RollbackTrans
    Exit Sub
    Resume
End Sub
Private Function Restaurar() As Boolean
Dim Data(1) As String, Log(1) As String
Dim i As Integer
Restaurar = False
On Error GoTo restarurar
    Set cnxtrans = New ADODB.Connection
    With cnxtrans
        .CursorLocation = adUseClient
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=TRANSFER;Data Source=" & VGParamSistem.Servidor
        .Open
    End With
    Set rstrans = New ADODB.Recordset
    rstrans.Open "Select name,filename From SysFiles ", cnxtrans, adOpenKeyset, adLockReadOnly
    rstrans.MoveFirst
    Data(0) = rstrans!Name: Data(1) = rstrans!FileName
    rstrans.MoveNext
    Log(0) = rstrans!Name: Log(1) = rstrans!FileName
Reintento:
    MsgBox "Se procede a Restaurar", vbInformation
    Set cnxtrans = Nothing
    Set rstrans = Nothing
    MsgBox "Se procede a Restaurar", vbInformation
    sqlcad = " RESTORE DATABASE TRANSFER " & _
             " FROM DISK='" & NombreArchivo & "'" & _
             " WITH RECOVERY,  " & _
             " Move '" & Trim$(Data(0)) & "' TO '" & Trim$(Data(1)) & "', " & _
             " Move '" & Trim$(Log(0)) & "' TO  '" & Trim$(Log(1)) & "' "
    VGGeneral.Execute sqlcad
    Restaurar = True
    MsgBox "Se Restauro Satisfactoriamente", vbInformation
    Exit Function
restarurar:
    If MsgBox("Desea Reintentar la Restauracion", vbQuestion + vbRetryCancel) = vbRetry Then
        GoTo Reintento
    End If
    Restaurar = False
    MsgBox err.Description, vbExclamation, "Error al restauar"
End Function

Private Sub Form_Load()
    DTPPerido.Month = CInt(VGParamSistem.Mesproceso)
    DTPPerido.Year = CInt(VGParamSistem.Anoproceso)
    Opt(1).Value = True
    Call LlenarCboSeries
End Sub

Private Sub Opt_Click(Index As Integer)
    Select Case Index
        Case 0
            CmdArchivo.Enabled = True
            cboSerie.Enabled = True
        Case 1
            CmdArchivo.Enabled = False
            cboSerie.Enabled = False
    End Select
End Sub

Sub LlenarCboSeries()
  cboSerie.AddItem "Serie 001"
  cboSerie.AddItem "Serie 002"
  cboSerie.AddItem "Serie 003"
  cboSerie.AddItem "Serie 004"
  cboSerie.AddItem "Serie 006"
  cboSerie.AddItem "Serie 008"
  cboSerie.AddItem "Serie 009"
  cboSerie.AddItem "Serie 011"
  cboSerie.AddItem "Serie 012"
  cboSerie.AddItem "Serie 013"
  
  SerieSubAsiento(0) = "0010"
  SerieSubAsiento(1) = "0009"
  SerieSubAsiento(2) = "0004"
  SerieSubAsiento(3) = "0001"
  SerieSubAsiento(4) = "0002"
  SerieSubAsiento(5) = "0008"
  SerieSubAsiento(6) = "0007"
  SerieSubAsiento(7) = "0003"
  SerieSubAsiento(8) = "0005"
  SerieSubAsiento(9) = "0006"
  cboSerie.ListIndex = -1

End Sub
