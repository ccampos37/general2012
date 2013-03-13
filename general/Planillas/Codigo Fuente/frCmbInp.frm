VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frCmbInp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar valores"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "frCmbInp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   2190
      TabIndex        =   1
      Top             =   2535
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   795
      TabIndex        =   3
      Top             =   2535
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adelantos"
      Height          =   2265
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4065
      Begin VB.CommandButton cmEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   2850
         TabIndex        =   5
         Top             =   1830
         Width           =   1020
      End
      Begin AplisetControlText.Aplitext xNuevoSaldo 
         Height          =   285
         Left            =   1545
         TabIndex        =   6
         Top             =   1860
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
         Redondear       =   -1  'True
      End
      Begin AplisetControlText.Aplitext xCambiar 
         Height          =   285
         Left            =   1545
         TabIndex        =   7
         Top             =   1245
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         Text            =   "0"
         Redondear       =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMontoOrigen 
         Height          =   285
         Left            =   1545
         TabIndex        =   8
         Top             =   915
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xDescripcion 
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   600
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto Original"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   975
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cambiar a"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1290
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Saldo"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   1905
         Width           =   930
      End
      Begin VB.Label xcodigo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3045
         TabIndex        =   4
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   315
         Width           =   840
      End
   End
End
Attribute VB_Name = "frCmbInp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMELIMINAR_CLICK()
    If UCase(Frame1.Caption) = "ADELANTOS" Then
        If MsgBox("REALMENTE DESEA QUITAR EL ADELANTO DEL CALCULO DE PLANILLA DE REMUNERACIONES QUE ESTÁ TRABAJANDO. AL QUITAR, EL MONTO DE ADELANTO ESTARÁ DISPONIBLE PARA SER DEBITADO SU SALDO EN OTRA PLANILLA. DESEA CONTINUAR CON EL PROCESO", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        DBSYSTEM.Execute "DELETE FROM  [##ADELANTOS" & VGL_COMPUTER & "]  WHERE CODIGO=" & xCodigo.Caption
        DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET ADELANTO=ADELANTO-" & xCambiar.Text & " WHERE CODTRAB='" & InputPl.xCodTrab.Tag & "'"
    Else
        If MsgBox("REALMENTE DESEA QUITAR EL PAGO DE CUENTA CORRIENTE. DESEA CONTINUAR CON EL PROCESO", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        DBSYSTEM.Execute "DELETE FROM  [##PAGOSCTACTE" & VGL_COMPUTER & "]  WHERE CODMOV=" & xCodigo.Caption
        If Left(InputPl.Lista.SelectedItem.Text, 1) = "I" Then
            DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET OTROSING=OTROSING-" & xCambiar.Text & " WHERE CODTRAB='" & InputPl.xCodTrab.Tag & "'"
        Else
            DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET OTROSEGR=OTROSEGR-" & Val(xCambiar.Text) & " WHERE CODTRAB='" & InputPl.xCodTrab.Tag & "'"
        End If
    End If
    InputPl.CALCULOTOTAL
    InputPl.REFRESCARTRAB
    Unload Me
End Sub

Private Sub COMMAND1_CLICK() 'ACEPTAR
    If Val(xCambiar.Text) >= Val(xMontoOrigen.Text) Then
        MsgBox "EL MONTO NUEVO DE REEMPLAZO NO ES CORRECTO. NO DEBE SER IGUAL O MAYOR QUE EL MONTO ORIGINAL", vbCritical
        Exit Sub
    End If
    If Val(xCambiar.Text) <= 0 Then
        MsgBox "ES MONTO NUEVO DE REEMPLAZO NO ES CORRECTO. NO PUEDE SER MENOR O IGUAL A CERO", vbCritical
        Exit Sub
    End If
    If UCase(Frame1.Caption) = "ADELANTOS" Then
        'ADELANTOS
        If MsgBox("LOS CAMBIOS NO SE PODRÁN DESHACER, PUES AFECTAN DIRECTAMENTE A LA BASE DE DATOS DEL SISTEMA DE PLANILLA. CONTINUAR CON LA ACCIÓN", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        DBSYSTEM.Execute "UPDATE  [##ADELANTOS" & VGL_COMPUTER & "]  SET MONTO=" & xCambiar.Text & " WHERE CODIGO=" & xCodigo.Caption
        DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET ADELANTO=ADELANTO-" & (Val(xMontoOrigen.Text) - Val(xCambiar.Text)) & " WHERE CODTRAB='" & InputPl.xCodTrab.Tag & "'"
        DBSYSTEM.Execute "UPDATE ADEL2000 SET MONTO=" & xCambiar.Text & " WHERE CODIGO=" & xCodigo.Caption
        DBSYSTEM.Execute "INSERT INTO ADEL2000 (CODTRAB,MES,FECHAING,MONTO,NUMBOL,NOMBOL) VALUES ('" & InputPl.xCodTrab.Tag & "'," & DateSQL(REGINPUT.MESACTIVO) & "," & DateSQL(Date) & "," & xNuevoSaldo.Text & ",0,0)"
    Else  'SI ES CARGO POR CUENTA CORRIENTE
        DBSYSTEM.Execute "UPDATE  [##PAGOSCTACTE" & VGL_COMPUTER & "]  SET CUOTA=" & xCambiar.Text & " WHERE CODMOV=" & xCodigo.Caption
        If Left(InputPl.Lista.SelectedItem.Text, 1) = "I" Then
            DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET OTROSING=OTROSING-" & (Val(xMontoOrigen.Text) - Val(xCambiar.Text)) & " WHERE CODTRAB='" & InputPl.xCodTrab.Tag & "'"
        Else
            DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET OTROSEGR=OTROSEGR-" & (Val(xMontoOrigen.Text) - Val(xCambiar.Text)) & " WHERE CODTRAB='" & InputPl.xCodTrab.Tag & "'"
        End If
    End If
    InputPl.CALCULOTOTAL
    InputPl.REFRESCARTRAB
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Select Case InputPl.Lista.SelectedItem.Tag
        Case 2 'ADELANTO DE PAGO
            xCodigo.Caption = InputPl.Lista.SelectedItem.Text
            xDescripcion.Text = InputPl.Lista.SelectedItem.SubItems(1)
            xMontoOrigen.Text = InputPl.Lista.SelectedItem.SubItems(2)
            xCambiar.Text = xMontoOrigen.Text
            xNuevoSaldo.Text = "0.00"
        Case 3 'CUENTAS CORRIENTES
            Label4.Visible = False
            xCodigo.Caption = Right(InputPl.Lista.SelectedItem.Text, Len(InputPl.Lista.SelectedItem.Text) - 2)
            xDescripcion.Text = InputPl.Lista.SelectedItem.SubItems(1)
            xMontoOrigen.Text = InputPl.Lista.SelectedItem.SubItems(2)
            xCambiar.Text = xMontoOrigen.Text
            xNuevoSaldo.Visible = False
            Frame1.Caption = "CUENTAS CORRIENTES"
    End Select
End Sub

Private Sub XCAMBIAR_CHANGE()
    xNuevoSaldo.Text = Val(xMontoOrigen.Text) - Val(xCambiar.Text)
End Sub

