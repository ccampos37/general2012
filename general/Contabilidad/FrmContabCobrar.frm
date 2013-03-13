VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmContabCobrar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de Datos y Generacion de Asientos de Cobranzas"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkRest 
      Caption         =   "Con Restauracion"
      Height          =   240
      Left            =   60
      TabIndex        =   10
      Top             =   1155
      Visible         =   0   'False
      Width           =   2865
   End
   Begin MSComDlg.CommonDialog cmdg_archivo 
      Left            =   6075
      Top             =   1695
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3420
      TabIndex        =   7
      Top             =   1875
      Width           =   1740
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1875
      Width           =   1740
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo de Proceso"
      Height          =   1005
      Left            =   3420
      TabIndex        =   5
      Top             =   90
      Width           =   3150
      Begin MSComCtl2.DTPicker DTPPerido 
         Height          =   315
         Left            =   1005
         TabIndex        =   8
         Top             =   420
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM - yyyy"
         Format          =   39845891
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
      Top             =   1470
      Width           =   375
   End
   Begin TextFer.TxFer TxArchivo 
      Height          =   360
      Left            =   45
      TabIndex        =   3
      Top             =   1455
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
         Caption         =   "Data Original"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   645
         Width           =   2520
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Desde Archivo"
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   300
         Width           =   2520
      End
   End
End
Attribute VB_Name = "FrmContabCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim sqlcad As String
Dim cnxtrans As ADODB.Connection
Dim rstrans As ADODB.Recordset
Dim NombreArchivo As String

Private Sub Form_Load()
    DTPPerido.Month = CInt(VGParamSistem.Mesproceso)
    DTPPerido.Year = CInt(VGParamSistem.Anoproceso)
    Opt(0).Enabled = False
    Opt(1).Value = True
End Sub

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
On Error GoTo Proceso
Screen.MousePointer = 11

paso1 = 0

Set rsparimpo = New ADODB.Recordset
rsparimpo.Open "ct_importarcobrar", VGCNx, adOpenKeyset, adLockReadOnly
If rsparimpo.RecordCount = 0 Then Exit Sub
If Opt(0).Value = True Then
    If ChkRest.Value = 1 Then
        If Not Restaurar Then
            Screen.MousePointer = 1
            Exit Sub
        End If
    End If
    BaseOrigen = "TRANSFER"
End If


Call DeleteDataPlanillaCobrar(rsparimpo!asientocodigo, rsparimpo!subasientocodigo, Format(Month(DTPPerido), "00"), Year(DTPPerido))
    VGGeneral.BeginTrans
    Set Comando = New ADODB.Command
    With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "cc_GeneraAsientoCobrarenLinea_pro"
        .ActiveConnection = VGGeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@BaseVenta") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@Asiento") = rsparimpo!asientocodigo
        .Parameters("@SubAsiento") = rsparimpo!subasientocodigo
        .Parameters("@Libro") = rsparimpo!Librocodigo
        .Parameters("@Mes") = Format(Month(DTPPerido), "00")
        .Parameters("@Ano") = Year(DTPPerido)
        .Parameters("@tipanal") = "002"
        .Parameters("@Compu") = VGcomputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@ajustedebe") = VGParametros.sistemactaajustedeb
        .Parameters("@ajustehaber") = VGParametros.sistemactaajustehab
        .Execute
    End With
    VGGeneral.CommitTrans
    Screen.MousePointer = 1
    MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Ventas"
    Unload Me
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

Private Sub Opt_Click(Index As Integer)
    Select Case Index
        Case 0
            CmdArchivo.Enabled = True
        Case 1
            CmdArchivo.Enabled = False
    End Select
End Sub

Sub DeleteDataPlanillaCobrar(Asiento As String, SubAsiento As String, mes As Integer, anno As String)
 Dim SQL As String
 
  SQL = "DELETE FROM ct_cabcomprob" & anno & " where asientocodigo='" & Asiento & "' AND "
  SQL = SQL & "subasientocodigo='" & SubAsiento & "' AND "
  SQL = SQL & "cabcomprobmes=" & mes & " and empresacodigo='" & VGParametros.empresacodigo & "'"
  VGCNx.Execute (SQL)

End Sub

