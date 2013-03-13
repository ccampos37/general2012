VERSION 5.00
Begin VB.Form frmTraEmi1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículo"
   ClientHeight    =   7368
   ClientLeft      =   1380
   ClientTop       =   1860
   ClientWidth     =   6132
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7368
   ScaleWidth      =   6132
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   105
      TabIndex        =   39
      Top             =   4740
      Width           =   5895
      Begin VB.TextBox txtCo1 
         Height          =   285
         Left            =   1320
         TabIndex        =   41
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtordfab 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   40
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Orden fab."
         Height          =   192
         Left            =   288
         TabIndex        =   42
         Top             =   288
         Width           =   744
      End
   End
   Begin VB.ComboBox cmbtipo 
      Height          =   315
      ItemData        =   "frmTraEmi1.frx":0000
      Left            =   1425
      List            =   "frmTraEmi1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   135
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   105
      TabIndex        =   27
      Top             =   3420
      Width           =   5895
      Begin VB.TextBox txtPIg 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblTNe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Left            =   1320
         TabIndex        =   35
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2160
         TabIndex        =   34
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Importe IGV."
         Height          =   195
         Left            =   3480
         TabIndex        =   33
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblIgv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Left            =   4440
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblTCo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Porcen. IGV."
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total Neto"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Compra"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   105
      TabIndex        =   20
      Top             =   2100
      Width           =   5895
      Begin VB.TextBox txtPDe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPUn 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblPNe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2160
         TabIndex        =   26
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Importe Dsct."
         Height          =   195
         Left            =   3360
         TabIndex        =   25
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblDes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Left            =   4440
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Precio Unit."
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Porcen. Dsct."
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Precio Neto"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2145
      Left            =   105
      TabIndex        =   10
      Top             =   -60
      Width           =   5895
      Begin VB.TextBox txtRef 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4830
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1740
         Width           =   855
      End
      Begin VB.TextBox txtURe 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3045
         TabIndex        =   2
         Top             =   1755
         Width           =   735
      End
      Begin VB.TextBox txtDes 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000015&
         Height          =   285
         Left            =   1320
         MaxLength       =   80
         TabIndex        =   7
         Top             =   975
         Width           =   4350
      End
      Begin VB.TextBox txtCan 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1365
         Width           =   1095
      End
      Begin VB.TextBox txtCod 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label lbltipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   285
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   195
         Left            =   3930
         TabIndex        =   19
         Top             =   1845
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Unid. Ref."
         Height          =   195
         Left            =   2280
         TabIndex        =   18
         Top             =   1815
         Width           =   720
      End
      Begin VB.Label lblFab 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4230
         TabIndex        =   17
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1425
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unidad"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1845
         Width           =   510
      End
      Begin VB.Label lblUnidad 
         AutoSize        =   -1  'True
         Caption         =   "Fabricante"
         Height          =   195
         Left            =   3240
         TabIndex        =   12
         Top             =   630
         Width           =   750
      End
      Begin VB.Label lblUni 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   1740
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   3504
      Picture         =   "frmTraEmi1.frx":0021
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6672
      Width           =   775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   1728
      Picture         =   "frmTraEmi1.frx":0463
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6624
      Width           =   775
   End
End
Attribute VB_Name = "frmTraEmi1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public activado As Boolean
Public cancelado As Boolean
Public Igv As Single
Public Tipo As String
Dim Mensaje As String

Private Sub cmdCancel_Click()
    cancelado = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtCod = "" Then
        Mensaje = "Debe ingresar Código de Artículo"
        MsgBox Mensaje, vbExclamation, "Error"
        txtCod.SetFocus
        Exit Sub
    Else
        If Not txtDes.Enabled Then
            If Not Existe(1, txtCod, "maeart", "acodigo", False) And Not Existe(1, txtCod, "Servicios", "ser_codigo", False) Then
                Mensaje = "Código de Artículo no válido"
                MsgBox Mensaje, vbExclamation, "Error"
                txtCod.SetFocus
                Exit Sub
            Else
                txtCod_KeyPress 13
                cmdOK.SetFocus
            End If
        End If
    End If
    
If Tipo = "B" Then
    If Val(txtCan) = 0 Then
        Mensaje = "Debe especificar Cantidad"
        MsgBox Mensaje, vbExclamation, "Error"
        txtCan.SetFocus
        Exit Sub
    End If
 End If
    
    If txtURe <> "" Then
        If Not txtRef.Enabled Then
            If Not Existe(1, txtURe, "tabunimed", "um_abrev", False) Then
                Mensaje = "Unidad de referencia no válida"
                MsgBox Mensaje, vbExclamation, "Error"
                txtURe.SetFocus
                Exit Sub
            Else
                txtURe_KeyPress 13
                cmdOK.SetFocus
            End If
        End If
        If Val(txtRef) = 0 Then
            Mensaje = "Debe especificar Orden de FabricacionccionReferencia"
            MsgBox Mensaje, vbExclamation, "Error"
            txtRef.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtPUn) = 0 Then
        Mensaje = "Debe especificar Precio Unitario"
        MsgBox Mensaje, vbExclamation, "Error"
        txtPUn.SetFocus
        Exit Sub
    End If
    
    cancelado = False
    txtCod.Enabled = True
    txtCod.SetFocus
    Me.Hide
End Sub

Private Sub Form_Activate()
    Igv = Val(txtPIg)
End Sub

Private Sub Form_Load()
   ' central Me
End Sub

Private Sub txtCan_Change()
    If Val(txtCan) = 0 Then
        txtURe = ""
        txtURe.Enabled = False
        txtPUn = "0.00"
        txtPUn.Enabled = False
    Else
        txtURe.Enabled = True
        txtPUn.Enabled = True
    End If
    Calculo_Automatico
End Sub

Private Sub txtCan_GotFocus()
    Enfoque txtCan
End Sub

Private Sub txtCan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtCan) > 0 Then
            txtURe.Enabled = True
            txtURe.SetFocus
        Else
            txtCan.SetFocus
        End If
    End If
    Reales_Positivos KeyAscii, txtCan
End Sub
Private Sub Reales_Positivos(k As Integer, t As TextBox)
Dim t1 As String
    k = Asc(UCase(Chr(k)))
    If k = 8 Then Exit Sub
    If k <> 45 And k <> 44 And k <> 32 And k <> 69 And k <> 43 Then
        t1 = Left(t, t.SelStart)
        t1 = t1 & Chr(k) & Right(t, Len(t) - Len(t1))
        If IsNumeric(t1) Then Exit Sub
    End If
    k = 0
    
End Sub

Private Sub txtCan_LostFocus()
    txtCan = Format(Val(txtCan), "0.00")
End Sub

Private Sub txtordfab_GotFocus()
    Enfoque txtordfab
End Sub

Private Sub txtordfab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCo1.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtco1_GotFocus()
    Enfoque txtCo1
End Sub

Private Sub txtCo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtCod_Change()
    If txtDes.Enabled Then
        lblFab = ""
        txtDes = ""
        txtDes.Enabled = False
        lblUni = ""
        txtCan = "0.00"
        txtCan.Enabled = False
    End If
End Sub

Private Sub txtCod_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    On Error GoTo Err
    Set Adodc2 = New ADODB.Recordset
    If frmEmisionOC.txtNSol <> "" Then
      strsql = "SELECT sccodigo,scdescri,scunidad FROM scd001 where scnumdoc='" & frmEmisionOC.txtNSol & "' "
      Adodc2.Open strsql, VGcnx, adOpenStatic, adLockReadOnly
    
      frmReferencia.Conectar Adodc2, strsql
      frmReferencia.Label1 = "Lista de Artículos"
      frmReferencia.Inicio
      frmReferencia.Show vbModal
      Adodc2.Close
    
     If vGUtil(1) <> "" Then
         txtCod = vGUtil(1)
         txtCod_KeyPress 13
     End If
  Else
      If cmbtipo.Text = "Bienes" Then
         strsql = "SELECT acodigo,adescri,aunidad FROM maeart "
         Adodc2.Open strsql, VGcnx, adOpenStatic, adLockReadOnly
    
        frmReferencia.Conectar Adodc2, strsql
        frmReferencia.Label1 = "Lista de Artículos"
        frmReferencia.Inicio
        frmReferencia.Show vbModal
        Adodc2.Close
    
       If vGUtil(1) <> "" Then
          txtCod = vGUtil(1)
          txtDes = vGUtil(2)
          txtCod_KeyPress 13
       End If
    ElseIf cmbtipo.Text = "Servicios" Then
         strsql = "SELECT ser_codigo,ser_descripcion FROM Servicios "
         Adodc2.Open strsql, VGcnx, adOpenStatic, adLockReadOnly
    
         frmReferencia.Conectar Adodc2, strsql
         frmReferencia.Label1 = "Lista de Servicios"
         frmReferencia.Inicio
         frmReferencia.Show vbModal
         Adodc2.Close
    
        If vGUtil(1) <> "" Then
           txtCod = vGUtil(1)
           txtDes = vGUtil(2)
           txtCod_KeyPress 13
        End If
    End If
 End If
 Exit Sub
Err:
  MsgBox Err.Description
End Sub

Private Sub txtCod_GotFocus()
    Enfoque txtCod
End Sub

Private Sub Enfoque(OBJ As Object)
  OBJ.SelStart = 0
  OBJ.SelLength = Len(OBJ)
End Sub

Private Sub txtCod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtCod_DblClick
End Sub

Public Function Existe(Tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
 Else
       If UCase(Tabla) = "PUNTO_VENTA" Then
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       End If
       If Trim(Cod2) <> "" Then
            cSL = cSL & " And  " & cCampo2 & " =  '" & SupCadSQL(Cod2) & "'"
       End If
       If Trim(Cod3) <> "" Then
            cSL = cSL & " And  " & cCampo3 & " =  '" & SupCadSQL(Cod3) & "'"
       End If
       If Trim(cCampo4) <> "" Then
            If Cod4 = True Then
                cSL = cSL & " And  " & cCampo4
            Else
                cSL = cSL & " And  " & Not cCampo4
            End If
        End If
        If Trim(Cod5) <> "" Then
            cSL = cSL & " And  " & cCampo5 & " =  '" & Cod5 & "'"
        End If
 End If
 
Select Case Tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Existe = True
Else
     Existe = False
End If
'csel1.Close
End Function

Public Function SupCadSQL(S As String) As String
 Dim Aux As String
 If Not IsNull(S) Then
     Aux = Replace(S, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function
Private Sub txtCod_KeyPress(KeyAscii As Integer)
 Dim adoreg As ADODB.Recordset
 Set adoreg = New ADODB.Recordset
 On Error GoTo Err
    If KeyAscii = 13 Then
        txtCod = Trim(txtCod)
        If txtCod <> "" Then
            If Not Existe(1, txtCod, "scd001", "sccodigo", False) And Not Existe(1, txtCod, "maeart", "acodigo", False) And Not Existe(1, txtCod, "servicios", "ser_codigo", False) Then
                Mensaje = "El Código de Artículo ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtCod.SetFocus
            Else
               If frmEmisionOC.txtNSol <> "" Then
                  adoreg.Open "Select sccantid,tipsol from scd001 where scnumdoc='" & frmEmisionOC.txtNSol & "' and sccodigo='" & txtCod & "'", VGcnx, adOpenStatic, adLockPessimistic
                 If Not adoreg.EOF Then
                     Tipo = adoreg("tipsol")
                  If Tipo = "B" Then
                       lblFab = Devolver_Dato(1, txtCod, "maeart", "acodigo", False, "acodigo2")
                       txtDes = Devolver_Dato(1, txtCod, "maeart", "acodigo", False, "adescri")
                       lblUni = Devolver_Dato(1, txtCod, "maeart", "acodigo", False, "aunidad")
                       txtCan.Enabled = True
                       txtCan = adoreg("sccantid")
                       txtCan.SetFocus
                    ElseIf Tipo = "S" Then
                       txtDes = Devolver_Dato(1, txtCod, "Servicios", "ser_codigo", False, "ser_descripcion")
                       txtPUn.Enabled = True
                       txtPUn.SetFocus
                   End If
                 End If
               Else
                  If cmbtipo.Text = "Bienes" Then
                     Tipo = "B"
                  ElseIf cmbtipo.Text = "Servicios" Then
                     Tipo = "S"
                  End If
                   If Tipo = "B" Then
                       lblFab = Devolver_Dato(1, txtCod, "maeart", "acodigo", False, "acodigo2")
                       txtDes = Devolver_Dato(1, txtCod, "maeart", "acodigo", False, "adescri")
                       lblUni = Devolver_Dato(1, txtCod, "maeart", "acodigo", False, "aunidad")
                       txtCan.Enabled = True
                       If txtCan = 0 Then
                         txtCan = 0
                       End If
                       txtCan.SetFocus
                    ElseIf Tipo = "S" Then
                       txtDes = Devolver_Dato(1, txtCod, "Servicios", "ser_codigo", False, "ser_descripcion")
                       txtPUn.Enabled = True
                       txtPUn.SetFocus
                   End If
               End If
              txtDes.Enabled = False
        End If
    Else
        txtCod.SetFocus
    End If
 Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
 Exit Sub
Err:
   MsgBox Err.Description
End Sub

Private Sub txtDes_GotFocus()
    Enfoque txtDes
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCan.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPDe_Change()
    Calculo_Automatico
End Sub

Private Sub txtPDe_GotFocus()
    Enfoque txtPDe
End Sub

Private Sub txtPDe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPIg.SetFocus
    End If
    Reales_Positivos KeyAscii, txtPDe
End Sub

Private Sub txtPDe_LostFocus()
    txtPDe = Format(Val(txtPDe), "0.00")
End Sub

Private Sub txtPIg_Change()
    Calculo_Automatico
End Sub

Private Sub txtPIg_GotFocus()
    Enfoque txtPIg
End Sub

Private Sub txtPIg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtordfab.SetFocus
    End If
    Reales_Positivos KeyAscii, txtPIg
End Sub

Private Sub txtPIg_LostFocus()
    txtPIg = Format(Val(txtPIg), "0.00")
End Sub

Private Sub txtPUn_Change()
    If Val(txtPUn) = 0 Then
        txtPDe = "0.00"
        txtPDe.Enabled = False
        txtPIg = Format(Igv, "0.00")
        txtPIg.Enabled = False
    Else
        txtPDe.Enabled = True
        txtPIg.Enabled = True
    End If
    Calculo_Automatico
End Sub

Private Sub txtPUn_GotFocus()
    Enfoque txtPUn
End Sub

Private Sub txtPUn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtPUn) > 0 Then
            txtPDe.SetFocus
        Else
            txtPUn.SetFocus
        End If
    End If
    Reales_Positivos KeyAscii, txtPUn
End Sub

Private Sub txtPUn_LostFocus()
    txtPUn = Format(Val(txtPUn), "0.00")
End Sub

Private Sub txtRef_Change()
    If Val(txtRef) = 0 Then
        If Me.ActiveControl.Name <> "txtURe" Then
            txtPUn = "0.00"
            txtPUn.Enabled = False
        End If
    Else
        txtPUn.Enabled = True
    End If
End Sub

Private Sub txtRef_GotFocus()
    Enfoque txtRef
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtRef) > 0 Then
            txtPUn.SetFocus
        Else
            txtRef.SetFocus
        End If
    End If
    Reales_Positivos KeyAscii, txtRef
End Sub

Private Sub txtRef_LostFocus()
    txtRef = Format(Val(txtRef), "0.00")
End Sub

Private Sub txtURe_Change()
    If txtURe <> "" Then
        txtPUn = "0.00"
        txtPUn.Enabled = False
    Else
        txtPUn.Enabled = True
    End If
    txtRef = ""
    txtRef.Enabled = False
    Calculo_Automatico
End Sub

Private Sub txtURe_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT um_abrev,um_nombre FROM tabunimed"
    Adodc2.Open strsql, VGcnx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Unidades de Medida"
    frmReferencia.Inicio
    frmReferencia.Show vbModal
    Adodc2.Close
    
    If vGUtil(1) <> "" Then
        txtURe = vGUtil(1)
        txtURe_KeyPress 13
    End If
End Sub

Private Sub txtURe_GotFocus()
    Enfoque txtURe
End Sub

Private Sub txtURe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtURe_DblClick
End Sub

Private Sub txtURe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtURe = Trim(txtURe)
        If txtURe <> "" Then
            If Not Existe(1, txtURe, "tabunimed", "um_abrev", False) Then
                Mensaje = "La Unidad de medida de Referencia no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtURe.SetFocus
            Else
                If Not txtRef.Enabled Then
                    txtRef = "0.00"
                    txtRef.Enabled = True
                End If
                txtRef.SetFocus
            End If
        Else
            txtPUn.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Sub Calculo_Automatico()
    If Not activado Then Exit Sub
    
    If Not txtRef.Enabled Then
        lblDes = Format(Val(txtPUn) * Val(txtPDe) / 100, "0.00")
        lblPNe = Format(Val(txtPUn) - Val(lblDes), "0.00")
        If Tipo = "S" Then
          lblTCo = Format(1 * Val(lblPNe), "0.00")
        Else
         lblTCo = Format(Val(txtCan) * Val(lblPNe), "0.00")
        End If
    Else
'        lblDes = Format((Val(txtPUn) * Val(txtRef) / Val(txtCan)) * Val(txtPDe) / 100, "0.00")
        lblDes = Format(Val(txtPUn) * Val(txtPDe) / 100, "0.00")
'        lblPNe = Format(Val(txtPUn) * Val(txtRef) / Val(txtCan) - Val(lblDes), "0.00")
        lblPNe = Format(Val(txtPUn) - Val(lblDes), "0.00")
        lblTCo = Format(Val(txtRef) * Val(lblPNe), "0.00")
    End If

    lblIgv = Format(Val(lblTCo) * Val(txtPIg) / 100, "0.00")
    lblTNe = Format(Val(lblTCo) + Val(lblIgv), "0.00")
End Sub
