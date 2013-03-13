VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportarDatosTesoreria 
   Caption         =   "Importar Datos Tesorería"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Generacion"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   150
      Width           =   2835
      Begin VB.OptionButton Opt 
         Caption         =   "Desde Archivo"
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   300
         Width           =   2520
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Servidor Local"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   9
         Top             =   645
         Width           =   2520
      End
   End
   Begin VB.CommandButton CmdArchivo 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   6690
      TabIndex        =   6
      Top             =   1620
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo de Proceso"
      Height          =   1005
      Left            =   3480
      TabIndex        =   3
      Top             =   135
      Width           =   3150
      Begin MSComCtl2.DTPicker DTPPerido 
         Height          =   315
         Left            =   1005
         TabIndex        =   4
         Top             =   420
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM - yyyy"
         Format          =   65011715
         CurrentDate     =   37656
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo :"
         Height          =   270
         Left            =   150
         TabIndex        =   5
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   1755
      TabIndex        =   2
      Top             =   2145
      Width           =   1740
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3480
      TabIndex        =   1
      Top             =   2145
      Width           =   1740
   End
   Begin VB.CheckBox ChkRest 
      Caption         =   "Con Restauracion"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   1245
      Visible         =   0   'False
      Width           =   2865
   End
   Begin MSComDlg.CommonDialog cmdg_archivo 
      Left            =   6675
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TextFer.TxFer TxArchivo 
      Height          =   360
      Left            =   105
      TabIndex        =   7
      Top             =   1575
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
End
Attribute VB_Name = "frmImportarDatosTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlcad As String
Dim cnxtrans As ADODB.Connection
Dim rstrans As ADODB.Recordset
Dim NombreArchivo As String
Dim m_opcion As String

Private Sub Form_Load()
    Me.Width = 7320
    Me.Height = 3015
    DTPPerido.Month = CInt(VGParamSistem.MesProceso)
    DTPPerido.Year = CInt(VGParamSistem.AnoProceso)
    Me.Caption = Me.Caption & " - " & m_opcion

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
             " Move '" & Trim(Data(0)) & "' TO '" & Trim(Data(1)) & "', " & _
             " Move '" & Trim(Log(0)) & "' TO  '" & Trim(Log(1)) & "' "
    VGgeneral.Execute sqlcad
    Restaurar = True
    MsgBox "Se Restauro Satisfactoriamente", vbInformation
    Exit Function
restarurar:
    If MsgBox("Desea Reintentar la Restauracion", vbQuestion + vbRetryCancel) = vbRetry Then
        GoTo Reintento
    End If
    Restaurar = False
    MsgBox Err.Description, vbExclamation, "Error al restauar"
End Function

Private Sub Opt_Click(Index As Integer)
  Select Case Index
      Case 0
          CmdArchivo.Enabled = True
      Case 1
          CmdArchivo.Enabled = False
  End Select
End Sub

Property Let Opcion(valor As String)
   m_opcion = valor
End Property

Private Sub cmdAceptar_Click()
  Select Case m_opcion
    Case "Ingreso":
       Call ProcesarIngresos
    
    Case "Egreso":
       Call ProcesarEgresos
    
    Case "Transferencia":
       Call ProcesarTransferencia
  
  End Select

End Sub

Private Sub ProcesarIngresos()
Dim rsparimpo As ADODB.Recordset
Dim BaseOrigen As String
Dim paso1 As Integer
On Error GoTo Proceso
Screen.MousePointer = 11
'Falta Procedimiento de Validaciones
'Que existan Asientos por Documentos
'Validar que las series tenga subasientos y cuentas correspondientes
'Que el arhivo a Importar cumpla con el Formato requerido
'Que los de el archivo a importar esten dentro de ese rango de fechas
'Que existan los parametros de la generacion de Asiento
'Que existan los parametros de la configuracion de documentos asientos
'Que existan los parametros de la configuracion de series, subasientos y documentos
'Tienen que estar las cuentas configuradas en el plan de cuentas
'Que los Documentos no esten generados en contabilidad y si es que hay una modificacion
' Corregir el Comprobante

'Aqui Se hace la restauracion del archivo a importar
'RESTORE DATABASE TRANSFER
'   FROM DISK='D:\Archivos Fernando\Documentos\EXPO_01012003_17012003'
'   WITH NORECOVERY,
'   Move 'MyNwind_data_1' TO 'D:\MyData\MyNwind_data_1.mdf',
'   Move 'MyNwind_data_2' TO 'D:\MyData\MyNwind_data_2.ndf'
paso1 = 0

Set rsparimpo = New ADODB.Recordset
rsparimpo.Open "ct_importartesoreria", VGcnx, adOpenKeyset, adLockReadOnly
rsparimpo.Filter = "tipooperacion='" & LCase(m_opcion) & "'"

If Opt(0).Value = True Then
    If ChkRest.Value = 1 Then
        If Not Restaurar Then
            Screen.MousePointer = 1
            Exit Sub
        End If
    End If
    BaseOrigen = "TRANSFER"
End If
If Opt(1).Value = True Then
    BaseOrigen = rsparimpo!Base
End If
'@Baseconta, @Baseventa, @Asiento, @SubAsiento, @Libro, @Mes, @Ano, @tipanal, @Compu, @Usuario

Call EliminarDatosTesoreria(VGParamSistem.BDEmpresa, rsparimpo!Asiento, rsparimpo!SubAsiento, rsparimpo!Libro)

Set Comando = New ADODB.Command
    VGgeneral.BeginTrans
    With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "te_GeneraAsientosTesoreria_pro"
        .ActiveConnection = VGgeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@BaseVenta") = BaseOrigen
        .Parameters("@Asiento") = rsparimpo!Asiento
        .Parameters("@SubAsiento") = rsparimpo!SubAsiento
        .Parameters("@Libro") = rsparimpo!Libro
        
        .Parameters("@Mes") = Format(Month(DTPPerido), "00")
        .Parameters("@Ano") = Year(DTPPerido)
        
        .Parameters("@tipanal") = "002"
        .Parameters("@Compu") = VGComputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@TipoMov") = UCase(Left(LCase(m_opcion), 1))
        .Execute
    End With
    VGgeneral.CommitTrans
    Screen.MousePointer = 1
    MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Ventas"
    Unload Me
    Exit Sub
Proceso:
    Screen.MousePointer = 1
    MsgBox Err.Description
    If paso = 1 Then VGgeneral.RollbackTrans
End Sub

Private Sub ProcesarEgresos()
Dim rsparimpo As ADODB.Recordset
Dim BaseOrigen As String
Dim paso1 As Integer
On Error GoTo Proceso
Screen.MousePointer = 11

paso1 = 0
Set rsparimpo = New ADODB.Recordset
If Opt(1).Value = True Then
    rsparimpo.Open "ct_importartesoreria", VGcnx, adOpenKeyset, adLockReadOnly
    rsparimpo.Filter = "tipooperacion='" & LCase(m_opcion) & "'"
    If ChkRest.Value = 1 Then
        If Not Restaurar Then
            Screen.MousePointer = 1
            Exit Sub
        End If
        BaseOrigen = "TRANSFER"
    End If
End If
If Opt(0).Value = True Then
    BaseOrigen = VGParamSistem.BDEmpresa
End If
'@Baseconta, @Baseventa, @Asiento, @SubAsiento, @Libro, @Mes, @Ano, @tipanal, @Compu, @Usuario

'Call EliminarDatosTesoreria(VGParamSistem.BDEmpresa, rsparimpo!Asiento, rsparimpo!SubAsiento, rsparimpo!Libro)

Set Comando = New ADODB.Command
    VGgeneral.BeginTrans
    With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "te_GeneraAsientosTesoreria_pro"
        .ActiveConnection = VGgeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseVenta") = BaseOrigen
        .Parameters("@Asiento") = "10"  'rsparimpo!Asiento
        .Parameters("@SubAsiento") = "0001" 'rsparimpo!SubAsiento
        .Parameters("@Libro") = "03" ' rsparimpo!Libro
        .Parameters("@Mes") = Format(Month(DTPPerido), "00")
        .Parameters("@Ano") = Year(DTPPerido)
        .Parameters("@tipanal") = "002"
        .Parameters("@Compu") = VGComputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
 '       .Parameters("@TipoMov") = "%" ' UCase(Left(LCase(m_opcion), 1))
        .Execute
    End With
    VGgeneral.CommitTrans
    Screen.MousePointer = 1
    MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Ventas"
    Unload Me
    Exit Sub
Proceso:
    Screen.MousePointer = 1
    MsgBox Err.Description
    VGgeneral.RollbackTrans
End Sub

Private Sub ProcesarTransferencia()
Dim rsparimpo As ADODB.Recordset
Dim BaseOrigen As String
Dim paso1 As Integer
On Error GoTo Proceso
Screen.MousePointer = 11

paso1 = 0
Set rsparimpo = New ADODB.Recordset
rsparimpo.Open "ct_importartesoreria", VGcnx, adOpenKeyset, adLockReadOnly
rsparimpo.Filter = "tipooperacion='" & LCase(m_opcion) & "'"

If Opt(0).Value = True Then
    If ChkRest.Value = 1 Then
        If Not Restaurar Then
            Screen.MousePointer = 1
            Exit Sub
        End If
    End If
    BaseOrigen = "TRANSFER"
End If
If Opt(1).Value = True Then
    BaseOrigen = rsparimpo!Base
End If

Call EliminarDatosTesoreria(VGParamSistem.BDEmpresa, rsparimpo!Asiento, rsparimpo!SubAsiento, rsparimpo!Libro)

Set Comando = New ADODB.Command
    VGgeneral.BeginTrans
    With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "te_GeneraAsientosTesoreriaTransf_pro"
        .ActiveConnection = VGgeneral
        .Parameters.Refresh
        .Parameters("@BaseVenta") = BaseOrigen
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@Asiento") = rsparimpo!Asiento
        .Parameters("@SubAsiento") = rsparimpo!SubAsiento
        .Parameters("@Libro") = rsparimpo!Libro
        .Parameters("@Mes") = Format(Month(DTPPerido), "00")
        .Parameters("@Ano") = Year(DTPPerido)
        .Parameters("@Compu") = VGComputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Execute
    End With
    VGgeneral.CommitTrans
    Screen.MousePointer = 1
    MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
    Unload Me
    Exit Sub
Proceso:
    Screen.MousePointer = 1
    MsgBox Err.Description
    If paso = 1 Then VGgeneral.RollbackTrans
End Sub

Sub EliminarDatosTesoreria(Base As String, Asiento As String, SubAsiento As String, Libro As String)
Dim cmdEliminar As New ADODB.Command

Set cmdEliminar = New ADODB.Command
    VGgeneral.BeginTrans
    With cmdEliminar
        .CommandType = adCmdStoredProc
        .CommandText = "ct_EliminaAsientosTesoreria_pro"
        .ActiveConnection = VGgeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@Asiento") = Asiento
        .Parameters("@SubAsiento") = SubAsiento
        .Parameters("@Libro") = Libro
        .Parameters("@Mes") = Format(Month(DTPPerido.Value), "00")
        .Parameters("@Ano") = Year(DTPPerido.Value)
        
        .Execute
    End With
    VGgeneral.CommitTrans
End Sub
