VERSION 5.00
Begin VB.Form FormConfiguracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuraciones"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5175
   Icon            =   "FormParEnt.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Por Defecto"
      Height          =   1236
      Left            =   72
      TabIndex        =   17
      Top             =   3384
      Width           =   5052
      Begin VB.OptionButton Opt2 
         Caption         =   "Almacen Suministro"
         Height          =   204
         Left            =   2472
         TabIndex        =   22
         Top             =   864
         Width           =   2028
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Almacen Ventas"
         Height          =   204
         Left            =   888
         TabIndex        =   21
         Top             =   864
         Value           =   -1  'True
         Width           =   2028
      End
      Begin VB.TextBox txtAlm 
         Height          =   285
         Left            =   936
         MaxLength       =   3
         TabIndex        =   19
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblAlm 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1392
         TabIndex        =   20
         Top             =   360
         Width           =   2532
      End
      Begin VB.Label Label3 
         Caption         =   "Almacen"
         Height          =   264
         Left            =   216
         TabIndex        =   18
         Top             =   396
         Width           =   1128
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3360
      Left            =   72
      TabIndex        =   7
      Top             =   0
      Width           =   5025
      Begin VB.CheckBox ChkLadrillos 
         Caption         =   "Opciones para Ladrilleras"
         Height          =   405
         Left            =   192
         TabIndex        =   23
         Top             =   1752
         Width           =   2436
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2196
         MaxLength       =   8
         TabIndex        =   13
         Top             =   2592
         Width           =   1704
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2196
         MaxLength       =   8
         TabIndex        =   12
         Top             =   3000
         Width           =   1692
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2232
         MaxLength       =   2
         TabIndex        =   11
         Top             =   684
         Width           =   864
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Subdiario Costo Ventas"
         Height          =   255
         Left            =   216
         TabIndex        =   10
         Top             =   684
         Width           =   2052
      End
      Begin VB.CommandButton Cmdhora 
         Caption         =   "Ajuste Horaria"
         Height          =   360
         Left            =   3492
         TabIndex        =   9
         Top             =   1392
         Width           =   1290
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   2208
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1404
         Width           =   900
      End
      Begin VB.CheckBox Chksub 
         Caption         =   "Subdiario Salidas"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   288
         Width           =   1584
      End
      Begin VB.CheckBox ChkAuto 
         Caption         =   "Numeración Automatica bloqueada"
         Height          =   375
         Left            =   204
         TabIndex        =   2
         Top             =   1020
         Width           =   3495
      End
      Begin VB.CheckBox ChkIASA 
         Caption         =   "Definir IASA"
         Height          =   405
         Left            =   204
         TabIndex        =   3
         Top             =   1368
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2244
         MaxLength       =   2
         TabIndex        =   1
         Top             =   288
         Width           =   864
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   144
         X2              =   4680
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label1 
         Caption         =   "Definir Cuenta Costo de Ventas "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   16
         Top             =   2304
         Width           =   3672
      End
      Begin VB.Label Label5 
         Caption         =   "(Debe)"
         Height          =   240
         Left            =   204
         TabIndex        =   15
         Top             =   2628
         Width           =   792
      End
      Begin VB.Label Label6 
         Caption         =   "(Haber)"
         Height          =   240
         Left            =   180
         TabIndex        =   14
         Top             =   2988
         Width           =   1116
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2844
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "FormConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoreg As ADODB.Recordset

Private Sub Check1_Click()
        Text3.Enabled = True
End Sub

Private Sub Cmdhora_Click()
Dim numero As Integer
If MsgBox("Esta opción es usada cuando se ha registrado en ventas" & Chr(13) & "movimientos de almacen y la  hora no era extendida." & Chr(13) & "Esta opción solo puede se usada con el consultor," & Chr(13) & "registar la fecha del proceso...Continua", vbYesNo, "Aviso") = vbYes Then
    Set adoreg = New ADODB.Recordset
    adoreg.Open "SELECT * FROM MOVALMCAB WHERE CACODMOV = 'FT'", VGCNx, adOpenDynamic, adLockOptimistic
    While Not adoreg.EOF
      If Val(Mid(adoreg("cahora"), 1, 2)) < 8 Then
          numero = Val(Mid(adoreg("cahora"), 1, 2)) + 12
          adoreg("cahora") = Trim(Str(numero) & Mid(adoreg("cahora"), 3, 6))
          adoreg.Update
      End If
      adoreg.MoveNext
    Wend
    adoreg.Close
End If
End Sub

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Command2_Click()
Dim rs As ADODB.Recordset
Dim rsql As String, csql As String

On Error GoTo Err
Text1 = Trim(Text1)
Text2 = Trim(Text2)
     If Trim(Text2) = "" And Chksub.Value = 1 Then
                MsgBox "Ingrese el subdiario", vbInformation, "Aviso"
                Text2.SetFocus
    End If
     If Trim(Text1) = "" And ChkIASA.Value = 1 Then
                MsgBox "Ingrese el Codigo de IASA", vbInformation, "Aviso"
                Text1.SetFocus
    End If
    rsql = "select * from configuracion"
    Set adoreg = New ADODB.Recordset
    adoreg.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
    If ChkIASA.Value = 1 And ChkAuto.Value = 1 And Chksub.Value = 1 Then
            If adoreg.EOF Then   'SUBDIARIO
                     csql = "insert into configuracion(conf_codigo,cod_iasa,cod_flagiasa,cod_bloqueo) values ('" & Text2 & "', '" & Text1 & "',true,true)"
             Else
                     csql = "update configuracion set      cod_flagiasa = -1 ,cod_bloqueo = -1 , conf_codigo = '" & Text2 & "' , cod_iasa ='" & Text1 & "' "
            End If
   ElseIf ChkIASA.Value = 1 And ChkAuto.Value = 1 Then
            
            If adoreg.EOF Then   'SUBDIARIO
                     csql = "insert into configuracion(cod_iasa,cod_flagiasa,cod_bloqueo) values ( '" & Text1 & "',true,true)"
             Else
                     csql = "update configuracion set   cod_flagiasa = -1  , cod_bloqueo = -1 ,  cod_iasa ='" & Text1 & "' "
            End If
   ElseIf ChkAuto.Value = 1 And Chksub.Value = 1 Then
            If adoreg.EOF Then   'SUBDIARIO
                     csql = "insert into configuracion(conf_codigo,cod_bloqueo) values ('" & Text2 & "', true)"
             Else
                     csql = "update configuracion set   cod_bloqueo = 1 , conf_codigo = '" & Text2 & "'  "
            End If
    ElseIf ChkIASA.Value = 1 And Chksub.Value = 1 Then
            If adoreg.EOF Then   'SUBDIARIO
                     csql = "insert into configuracion(conf_codigo,cod_iasa,cod_flagiasa) values ('" & Text2 & "', '" & Text1 & "',true)"
             Else
                     csql = "update configuracion set conf_codigo = '" & Text2 & "' , cod_iasa ='" & Text1 & "' ,cod_flagiasa = -1  "
            End If
    ElseIf ChkIASA.Value = 1 Then
            If adoreg.EOF Then   'SUBDIARIO
                     csql = "insert into configuracion(cod_iasa,cod_flagiasa) values ( '" & Text1 & "',true)"
             Else
                     csql = "update configuracion set cod_iasa ='" & Text1 & "' , cod_flagiasa = -1  "
            End If
    ElseIf Chksub.Value = 1 Then
            If adoreg.EOF Then   'SUBDIARIO
                     csql = "insert into configuracion(conf_codigo) values ( '" & Text2 & "')"
             Else
                     csql = "update configuracion set conf_codigo = '" & Text2 & "'   "
            End If
     ElseIf ChkAuto.Value = 1 Then
            
            If adoreg.EOF Then
                     csql = "insert into configuracion(cod_bloqueo) values ( true)"
             Else
                     csql = "update configuracion set  cod_bloqueo = FALSE  "
            End If
   End If
   
   
   If csql <> "" Then VGCNx.Execute csql
   
   If ChkAuto.Value Then
      VGAutomatico = True
   Else
      csql = "update configuracion set  cod_bloqueo = 0  "
      VGAutomatico = False
      VGCNx.Execute csql
   End If
   
   
   csql = "Update configuracion set cosven_debe='" & Text5 & "',cosven_habe='" & Text4 & "',conf_codigoIng='" & IIf(Check1.Value = 1, Text3, "") & "',conf_codigo='" & IIf(Chksub.Value = 1, Text2, "") & "',cod_flagiasa =" & IIf(ChkIASA.Value = 1, 1, 0) & ",cod_iasa='" & Text1 & "',alma_defa='" & txtAlm & "',tipo_alma='" & IIf(opt1.Value = True, "V", "S") & "',Ladrillera='" & IIf(ChkLadrillos.Value = 1, "S", "N") & "'"
   VGCNx.Execute csql
   
   VGTip_Alma = IIf(opt1.Value = True, "V", "S")
   
   VGIASA = Text1
   Command1_Click
  Exit Sub
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
   MsgBox "Actulize Base de Datos........", vbInformation, "Aviso"

End Sub



Private Sub ChkAuto_Click()
 If ChkAuto.Value = 1 Then
   ' FormRegistro.Text4.Enabled = False
   
 End If
End Sub

Private Sub ChkIASA_Click()
  If ChkIASA.Value = 1 Then
      Text1.Enabled = True
     
      'MDIPrincipal.mnu_repIASA.Visible = True
      'MDIPrincipal.mnu_guiaIngIasa.Visible = True
      'MDIPrincipal.mnu_recepcion.Visible = True
  Else
      'MDIPrincipal.mnu_repIASA.Visible = False
      'MDIPrincipal.mnu_guiaIngIasa.Visible = False
     ' MDIPrincipal.mnu_recepcion.Visible = False
  End If
End Sub

Private Sub Chksub_Click()
   Text2.Enabled = True
End Sub

Private Sub Form_Load()
Dim rsql As String
On Local Error GoTo ERRAR
    central Me

    rsql = "select * from configuracion"
    Set adoreg = New ADODB.Recordset
    adoreg.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
    If Not adoreg.EOF Then
       Text2 = IIf(IsNull(adoreg("conf_codigo")), "", adoreg("conf_codigo"))
       Chksub = IIf(Text2 <> "", 1, 0)
       Text2.Enabled = IIf(Chksub.Value = 1, True, False)
       Text1 = IIf(IsNull(adoreg("cod_iasa")), "", adoreg("cod_iasa"))
       ChkIASA = IIf(Text1 <> "", 1, 0)
       Text2.Enabled = IIf(ChkIASA.Value = 1, True, False)
       ChkAuto = IIf(adoreg("cod_bloqueo"), 1, 0)
       
       Text3 = IIf(IsNull(adoreg("conf_codigoIng")), "", adoreg("conf_codigoIng"))
       Check1.Value = IIf(Text3 <> "", 1, 0)
       Text3.Enabled = IIf(Check1.Value = 1, True, False)
       Text5 = cNull(adoreg!cosven_debe)
       Text4 = cNull(adoreg!cosven_habe)
       txtAlm = cNull(adoreg!Alma_defa)
       If adoreg!tipo_Alma = "V" Then 'almacen ventas
          opt1.Value = True
          opt2.Value = False
       Else
          opt1.Value = False
          opt2.Value = True
       End If
       
       ChkLadrillos.Value = IIf(cNull(adoreg!Ladrillera) = "S", 1, 0)
       lblAlm = Devolver_Dato(1, txtAlm, "tabalm", "taalma", False, "tadescri")
    Else
       Text1.Enabled = False
       Text2.Enabled = False
    End If
    adoreg.Close
    
Exit Sub
ERRAR:
     MsgBox "Error en Base de DAtos Actualize Base de Datos en Herramientas", vbCritical, "Verifique "
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command2.SetFocus
End Sub



Private Sub Text2_DblClick()
Dim adoreg3 As ADODB.Recordset
Set adoreg3 = New ADODB.Recordset
 

If UCase(Dir$(cRuta4)) = VGNameCont & ".MDB" Then 'vGBdContabilidad
         adoreg3.Open "SELECT SUBDIAR_CODIGO,SUBDIAR_DESCRIPCION FROM SUBDIARIOS", VGcnxCT, adOpenStatic
        frmReferencia.Conectar adoreg3, "SELECT SUBDIAR_CODIGO,SUBDIAR_DESCRIPCION FROM SUBDIARIOS"
        frmReferencia.Label1.Caption = "Subdiarios"
        frmReferencia.Show vbModal
        adoreg3.Close
        If vGUtil(1) <> "" Then
            Text2 = vGUtil(1)
        End If
End If
End Sub

Private Sub Text3_DblClick()
Dim adoreg3 As ADODB.Recordset
Set adoreg3 = New ADODB.Recordset
 
If UCase(Dir$(cRuta4)) = VGNameCont & ".MDB" Then 'vGBdContabilidad
         adoreg3.Open "SELECT SUBDIAR_CODIGO,SUBDIAR_DESCRIPCION FROM SUBDIARIOS", VGcnxCT, adOpenStatic
        frmReferencia.Conectar adoreg3, "SELECT SUBDIAR_CODIGO,SUBDIAR_DESCRIPCION FROM SUBDIARIOS"
        frmReferencia.Label1.Caption = "Subdiarios"
        frmReferencia.Show vbModal
        adoreg3.Close
        If vGUtil(1) <> "" Then
            Text3 = vGUtil(1)
        End If
End If
End Sub

Private Sub Text4_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Dim cBase As String
cBase = cRuta4
If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
        Adodc2.Open "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional", VGcnxCT, adOpenStatic
        frmReferencia.Conectar Adodc2, "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional"
        frmReferencia.Label1.Caption = "Plan de Cuenta Nacional"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text4.text = (vGUtil(1))
        End If
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text4_DblClick
End Sub

Private Sub Text5_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Dim cBase As String
cBase = cRuta4
If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
        Adodc2.Open "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional", VGcnxCT, adOpenStatic
        frmReferencia.Conectar Adodc2, "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional"
        frmReferencia.Label1.Caption = "Plan de Cuenta Nacional"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text5.text = (vGUtil(1))
        End If
End If

End Sub

Private Sub txtAlm_Change()
    If lblAlm <> "" Then lblAlm = ""
End Sub
Private Sub txtAlm_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT taalma,tadescri FROM tabalm"
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Lista de Almacenes"
    'frmReferencia.i
    frmReferencia.Show vbModal
    Adodc2.Close
    
    If vGUtil(1) <> "" Then
        txtAlm = vGUtil(1)
        lblAlm = vGUtil(2)
        Command2.SetFocus
    End If
End Sub
    
Private Sub txtAlm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then txtAlm_DblClick
End Sub
Private Sub txtAlm_KeyPress(KeyAscii As Integer)
    Dim Mensaje As String
    If KeyAscii = 13 Then
        txtAlm = Trim(txtAlm)
        If txtAlm <> "" Then
            If Not Existe(1, txtAlm, "tabalm", "taalma", False) Then
                Mensaje = "El Código de Almacén ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtAlm.SetFocus
            Else
                lblAlm = Devolver_Dato(1, txtAlm, "tabalm", "taalma", False, "tadescri")
                Command2.SetFocus
            End If
        Else
            Command2.SetFocus
        End If
    End If
    'Enteros_Positivos KeyAscii, txtAlm
End Sub

